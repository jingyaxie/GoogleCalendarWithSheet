/**
 * Google Apps Script: 同步课程信息到日历并发送邮件
 * 
 * 功能：
 * 1. 从Google表格读取课程信息
 * 2. 发送邮件通知给老师和学生
 * 3. 创建日历事件到老师和学生的日历
 * 4. 在隐藏sheet中记录处理状态
 */

// ==================== 配置常量 ====================
const CONFIG = {
  // 主表名称（根据实际情况修改）
  MAIN_SHEET_NAME: '课程安排',
  
  // 隐藏状态表名称
  STATUS_SHEET_NAME: '_StatusLog',
  
  // 时区设置
  TIMEZONE: 'Asia/Shanghai',
  
  // 邮件模板
  EMAIL_TEMPLATE: {
    subject: '课程通知：{courseTitle}',
    body: `
      <html>
        <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
          <h2 style="color: #4CAF50;">课程通知</h2>
          <p>您好 {recipientName}，</p>
          <p>这是一封关于即将到来的课程通知：</p>
          <div style="background-color: #f5f5f5; padding: 15px; border-radius: 5px; margin: 20px 0;">
            <p><strong>课程主题：</strong>{courseTitle}</p>
            <p><strong>日期：</strong>{courseDate}</p>
            <p><strong>时间：</strong>{startTime} - {endTime}</p>
            <p><strong>老师：</strong>{teacherName}</p>
            <p><strong>学生：</strong>{studentName}</p>
          </div>
          <p>课程事件已添加到您的日历中，请及时查看。</p>
          <p>如有任何问题，请及时联系。</p>
          <p style="margin-top: 30px; color: #666; font-size: 12px;">此邮件由系统自动发送，请勿回复。</p>
        </body>
      </html>
    `
  }
};

// ==================== 主函数 ====================

/**
 * 主执行函数 - 处理所有课程记录
 */
function main() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // 确保隐藏状态表存在
    ensureStatusSheet(spreadsheet);
    
    // 读取主表数据
    const mainSheet = spreadsheet.getSheetByName(CONFIG.MAIN_SHEET_NAME);
    if (!mainSheet) {
      throw new Error(`找不到主表: ${CONFIG.MAIN_SHEET_NAME}`);
    }
    
    // 确保正式表有"记录ID"列
    ensureRecordIdColumn(mainSheet);
    
    const courses = readCourseData(mainSheet);
    Logger.log(`读取到 ${courses.length} 条课程记录`);
    
    // 读取已处理状态（在同步之前读取，以便检测被删除的记录）
    const statusSheet = spreadsheet.getSheetByName(CONFIG.STATUS_SHEET_NAME);
    const processedRecords = readProcessedStatus(statusSheet);
    
    // 检测被删除的记录（在同步状态表之前检测，避免状态表被删除后无法检测）
    const deletedRecords = findDeletedRecords(courses, processedRecords, statusSheet);
    if (deletedRecords.length > 0) {
      Logger.log(`检测到 ${deletedRecords.length} 条被删除的记录，将取消课程`);
      for (const deletedRecord of deletedRecords) {
        try {
          cancelCourse(deletedRecord, statusSheet);
          Logger.log(`取消课程成功: ${deletedRecord.lessonNumber} - ${deletedRecord.date}`);
        } catch (error) {
          Logger.log(`取消课程失败: ${deletedRecord.lessonNumber} - ${error.message}`);
        }
      }
    }
    
    // 同步状态表，确保和正式表一一对应（在检测被删除记录之后）
    syncStatusSheet(statusSheet, courses.length);
    
    // 重新读取已处理状态（同步后重新读取）
    const processedRecordsAfterSync = readProcessedStatus(statusSheet);
    
    // 为每条课程记录分配或获取记录ID，并更新正式表
    assignRecordIds(courses, processedRecordsAfterSync, statusSheet, mainSheet);
    
    // 计算每条课程的token并判断是否需要处理
    const toProcess = courses.filter(course => {
      // 优先通过记录ID查找，如果没有记录ID，则通过key查找（向后兼容）
      let existingRecord = null;
      if (course.recordId) {
        existingRecord = processedRecords.byId.get(course.recordId);
      }
      if (!existingRecord) {
        const key = `${course.lessonNumber}_${course.date}`;
        existingRecord = processedRecords.byKey.get(key);
      }
      
      if (!existingRecord) {
        // 新记录，需要处理
        // 检查是否有相同课次但不同日期的旧记录（日期变化）
        const oldRecords = findOldRecordsByLessonNumber(statusSheet, course.lessonNumber, course.date);
        if (oldRecords.length > 0) {
          Logger.log(`检测到日期变化: ${course.lessonNumber}，将在处理时删除旧日期的日历事件`);
          // 标记需要删除的旧记录，在processCourse中处理（因为需要日历ID）
          course._oldRecords = oldRecords;
        }
        return true;
      }
      
      // 计算当前记录的token
      const currentToken = calculateCourseToken(course);
      const existingToken = existingRecord.token || '';
      
      // 如果token不同，说明关键信息有变化，需要更新
      if (currentToken !== existingToken) {
        Logger.log(`检测到关键信息变化: ${course.lessonNumber} (旧token: ${existingToken}, 新token: ${currentToken})`);
        return true;
      }
      
      // token相同，说明关键信息没有变化
      // 检查是否已有日历事件ID，如果有则完全跳过（不发送邮件也不更新日历）
      if (existingRecord.teacherEventId || existingRecord.studentEventId) {
        Logger.log(`跳过处理（token相同且已有日历事件）: ${course.lessonNumber}`);
        return false;
      }
      
      // token相同但没有日历事件ID，可能是之前创建失败，需要重试
      // 但只有在状态不是已完成时才处理
      if (existingRecord.status !== '已完成') {
        Logger.log(`重试处理（token相同但之前失败）: ${course.lessonNumber}`);
        return true;
      }
      
      // token相同且已完成，跳过
      return false;
    });
    
    Logger.log(`需要处理 ${toProcess.length} 条记录`);
    
    // 处理每条记录
    const results = [];
    for (const course of toProcess) {
      try {
        const result = processCourse(course, statusSheet);
        results.push(result);
        Logger.log(`处理完成: ${course.lessonNumber} - ${result.status}`);
      } catch (error) {
        Logger.log(`处理失败: ${course.lessonNumber} - ${error.message}`);
        results.push({
          course: course,
          status: '失败',
          error: error.message
        });
      }
    }
    
    // 输出处理结果
    Logger.log('=== 处理结果汇总 ===');
    const successCount = results.filter(r => r.status === '已完成').length;
    const failCount = results.length - successCount;
    Logger.log(`成功: ${successCount}, 失败: ${failCount}`);
    
    return {
      total: results.length,
      success: successCount,
      failed: failCount,
      results: results
    };
    
  } catch (error) {
    Logger.log(`主函数执行错误: ${error.message}`);
    Logger.log(error.stack);
    throw error;
  }
}

/**
 * 处理单条课程记录
 */
function processCourse(course, statusSheet) {
  const result = {
    course: course,
    teacherEmail: { sent: false, eventId: null, error: null },
    studentEmail: { sent: false, eventId: null, error: null },
    status: '处理中'
  };
  
  try {
    // 如果有旧记录（日期变化），先删除旧日期的日历事件
    if (course._oldRecords && course._oldRecords.length > 0) {
      for (const oldRecord of course._oldRecords) {
        // 尝试删除老师日历事件（使用旧记录中的日历ID）
        if (oldRecord.teacherEventId) {
          try {
            if (oldRecord.teacherCalendarId) {
              // 如果有日历ID，直接删除
              deleteCalendarEvent(oldRecord.teacherCalendarId, oldRecord.teacherEventId);
              Logger.log(`删除旧老师日历事件成功: ${oldRecord.teacherEventId} (日历: ${oldRecord.teacherCalendarId})`);
            } else {
              // 如果没有日历ID，尝试通过事件ID删除（遍历所有日历）
              deleteCalendarEventById(oldRecord.teacherEventId);
              Logger.log(`删除旧老师日历事件成功: ${oldRecord.teacherEventId}`);
            }
          } catch (error) {
            Logger.log(`删除旧老师日历事件失败: ${oldRecord.teacherEventId} - ${error.message}`);
          }
        }
        
        // 尝试删除学生日历事件（使用旧记录中的日历ID）
        if (oldRecord.studentEventId) {
          try {
            if (oldRecord.studentCalendarId) {
              // 如果有日历ID，直接删除
              deleteCalendarEvent(oldRecord.studentCalendarId, oldRecord.studentEventId);
              Logger.log(`删除旧学生日历事件成功: ${oldRecord.studentEventId} (日历: ${oldRecord.studentCalendarId})`);
            } else {
              // 如果没有日历ID，尝试通过事件ID删除（遍历所有日历）
              deleteCalendarEventById(oldRecord.studentEventId);
              Logger.log(`删除旧学生日历事件成功: ${oldRecord.studentEventId}`);
            }
          } catch (error) {
            Logger.log(`删除旧学生日历事件失败: ${oldRecord.studentEventId} - ${error.message}`);
          }
        }
      }
      
      // 删除旧状态记录
      deleteOldStatusRecords(statusSheet, course._oldRecords);
    }
    
    // 获取已有的事件ID和token信息
    const existingInfo = getExistingEventIds(statusSheet, course);
    
    // 判断是否需要重新发送邮件（关键信息有变化时）
    const needsResendEmail = existingInfo.hasChanges;
    
    // 1. 发送老师邮件（仅在关键信息变化时发送）
    if (needsResendEmail) {
      try {
        sendCourseEmail(
          course.teacherEmail,
          course.teacherName,
          course,
          course.studentName
        );
        result.teacherEmail.sent = true;
        Logger.log(`老师邮件发送成功: ${course.teacherEmail}`);
      } catch (error) {
        result.teacherEmail.error = error.message;
        Logger.log(`老师邮件发送失败: ${error.message}`);
      }
    } else {
      Logger.log(`老师邮件跳过（关键信息未变化）: ${course.teacherEmail}`);
    }
    
    // 2. 创建或更新老师日历事件（仅在关键信息有变化或没有事件ID时）
    if (existingInfo.hasChanges || !existingInfo.teacherEventId) {
      try {
        const teacherEventId = createOrUpdateCalendarEvent(
          course.teacherCalendarId,
          course,
          existingInfo.teacherEventId
        );
        result.teacherEmail.eventId = teacherEventId;
        if (existingInfo.teacherEventId && existingInfo.hasChanges) {
          Logger.log(`老师日历事件更新成功: ${teacherEventId}`);
        } else if (existingInfo.teacherEventId) {
          Logger.log(`老师日历事件保持不变: ${teacherEventId}`);
        } else {
          Logger.log(`老师日历事件创建成功: ${teacherEventId}`);
        }
      } catch (error) {
        result.teacherEmail.error = error.message;
        Logger.log(`老师日历事件处理失败: ${error.message}`);
      }
    } else {
      // token相同且已有事件ID，跳过更新
      result.teacherEmail.eventId = existingInfo.teacherEventId;
      Logger.log(`老师日历事件跳过（token相同且已有事件）: ${existingInfo.teacherEventId}`);
    }
    
    // 3. 发送学生邮件（仅在关键信息变化时发送）
    if (needsResendEmail) {
      try {
        sendCourseEmail(
          course.studentEmail,
          course.studentName,
          course,
          course.teacherName
        );
        result.studentEmail.sent = true;
        Logger.log(`学生邮件发送成功: ${course.studentEmail}`);
      } catch (error) {
        result.studentEmail.error = error.message;
        Logger.log(`学生邮件发送失败: ${error.message}`);
      }
    } else {
      Logger.log(`学生邮件跳过（关键信息未变化）: ${course.studentEmail}`);
    }
    
    // 4. 创建或更新学生日历事件（仅在关键信息有变化或没有事件ID时）
    if (existingInfo.hasChanges || !existingInfo.studentEventId) {
      try {
        const studentEventId = createOrUpdateCalendarEvent(
          course.studentCalendarId,
          course,
          existingInfo.studentEventId
        );
        result.studentEmail.eventId = studentEventId;
        if (existingInfo.studentEventId && existingInfo.hasChanges) {
          Logger.log(`学生日历事件更新成功: ${studentEventId}`);
        } else if (existingInfo.studentEventId) {
          Logger.log(`学生日历事件保持不变: ${studentEventId}`);
        } else {
          Logger.log(`学生日历事件创建成功: ${studentEventId}`);
        }
      } catch (error) {
        result.studentEmail.error = error.message;
        Logger.log(`学生日历事件处理失败: ${error.message}`);
      }
    } else {
      // token相同且已有事件ID，跳过更新
      result.studentEmail.eventId = existingInfo.studentEventId;
      Logger.log(`学生日历事件跳过（token相同且已有事件）: ${existingInfo.studentEventId}`);
    }
    
    // 5. 判断整体状态
    // 如果邮件跳过（因为token没变化），不应该影响成功判断
    // 只要日历事件创建成功，就算成功
    const teacherSuccess = result.teacherEmail.eventId && !result.teacherEmail.error;
    const studentSuccess = result.studentEmail.eventId && !result.studentEmail.error;
    
    if (teacherSuccess && studentSuccess) {
      result.status = '已完成';
    } else if (teacherSuccess || studentSuccess) {
      result.status = '部分失败';
    } else {
      result.status = '失败';
    }
    
    // 6. 记录状态到隐藏sheet
    updateStatusRecord(statusSheet, course, result);
    
    return result;
    
  } catch (error) {
    result.status = '失败';
    result.error = error.message;
    updateStatusRecord(statusSheet, course, result);
    throw error;
  }
}

// ==================== 数据读取模块 ====================

/**
 * 读取课程数据
 */
function readCourseData(sheet) {
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  if (values.length < 2) {
    return [];
  }
  
  // 表头行（第1行，索引0）
  const headers = values[0];
  const headerMap = {};
  headers.forEach((header, index) => {
    headerMap[header.trim()] = index;
  });
  
  // 数据行（从第2行开始，索引1）
  const courses = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    
    // 跳过空行
    if (!row[0] || !row[headerMap['日期']]) {
      continue;
    }
    
    try {
      const course = {
        lessonNumber: row[headerMap['课次']] || '',
        date: row[headerMap['日期']] || '',
        courseTitle: row[headerMap['课程内容/主题']] || '',
        teacherName: row[headerMap['老师']] || '',
        teacherEmail: row[headerMap['老师邮箱']] || '',
        studentName: row[headerMap['学生']] || '',
        studentEmail: row[headerMap['学生邮箱']] || '',
        startTime: row[headerMap['开始时间']] || '',
        endTime: row[headerMap['结束时间']] || '',
        teacherCalendarId: row[headerMap['老师日历授权ID']] || row[headerMap['老师邮箱']] || '',
        studentCalendarId: row[headerMap['学生日历授权ID']] || row[headerMap['学生邮箱']] || '',
        rowIndex: i + 1 // 记录行号（正式表的行号，从1开始，包含表头），用于和状态表一一对应
      };
      
      // 获取记录ID（如果正式表有"记录ID"列，使用它）
      if (headerMap['记录ID'] !== undefined) {
        course.recordId = row[headerMap['记录ID']] || '';
      } else {
        course.recordId = ''; // 稍后从状态表获取或生成
      }
      
      // 记录记录ID列的索引（用于后续更新）
      course.recordIdColumnIndex = headerMap['记录ID'];
      
      // 计算token
      course.token = calculateCourseToken(course);
      
      // 验证必要字段
      if (!course.date || !course.teacherEmail || !course.studentEmail) {
        Logger.log(`跳过无效记录（第${i+1}行）: 缺少必要字段`);
        continue;
      }
      
      courses.push(course);
    } catch (error) {
      Logger.log(`解析第${i+1}行数据时出错: ${error.message}`);
      continue;
    }
  }
  
  return courses;
}

/**
 * 读取已处理状态（通过记录ID或行号索引，和正式表一一对应）
 */
function readProcessedStatus(statusSheet) {
  const processedMap = new Map();
  const processedMapById = new Map(); // 通过记录ID索引
  
  if (!statusSheet || statusSheet.getLastRow() < 2) {
    return { byKey: processedMap, byId: processedMapById };
  }
  
  const dataRange = statusSheet.getDataRange();
  const values = dataRange.getValues();
  
  // 从第2行开始读取（第1行为表头）
  // 列索引：0=记录ID, 1=课次, 2=日期, 3=Token, 6=老师日历ID, 7=老师日历事件ID, 11=学生日历ID, 12=学生日历事件ID, 14=处理状态
  // 状态表的第i行对应正式表的第i行（都有表头）
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    // 如果课次和日期都为空，跳过（空行）
    if (!row[1] && !row[2]) {
      continue;
    }
    
    const recordId = row[0] || ''; // 记录ID
    const key = `${row[1]}_${row[2]}`; // 课次_日期（向后兼容）
    
    const record = {
      recordId: recordId, // 记录ID
      lessonNumber: row[1],
      date: row[2],
      token: row[3] || '', // Token（关键信息哈希）
      teacherCalendarId: row[6] || '', // 老师日历ID（用于删除事件）
      teacherEventId: row[7] || '', // 老师日历事件ID（向后兼容：如果新格式没有，尝试旧格式）
      studentCalendarId: row[11] || '', // 学生日历ID（用于删除事件）
      studentEventId: row[12] || '', // 学生日历事件ID（向后兼容：如果新格式没有，尝试旧格式）
      status: row[14] || row[12] || '', // 处理状态（向后兼容）
      rowIndex: i + 1 // 状态表的行号（从1开始，包含表头）
    };
    
    // 向后兼容：如果新格式没有事件ID，尝试从旧格式读取
    if (!record.teacherEventId && row[5]) {
      record.teacherEventId = row[5]; // 旧格式：老师日历事件ID在第5列
    }
    if (!record.studentEventId && row[9]) {
      record.studentEventId = row[9]; // 旧格式：学生日历事件ID在第9列
    }
    
    // 通过key索引（向后兼容）
    processedMap.set(key, record);
    
    // 通过记录ID索引（优先使用）
    if (recordId) {
      processedMapById.set(recordId, record);
    }
  }
  
  return { byKey: processedMap, byId: processedMapById };
}

/**
 * 确保正式表有"记录ID"列
 */
function ensureRecordIdColumn(mainSheet) {
  const headers = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
  const hasRecordIdColumn = headers.some(header => header.trim() === '记录ID');
  
  if (!hasRecordIdColumn) {
    // 在最后一列添加"记录ID"列
    const lastColumn = mainSheet.getLastColumn();
    const newColumnIndex = lastColumn + 1;
    mainSheet.getRange(1, newColumnIndex).setValue('记录ID');
    Logger.log(`在正式表添加"记录ID"列: 第${newColumnIndex}列`);
  }
}

/**
 * 为课程记录分配或获取记录ID，并更新正式表
 */
function assignRecordIds(courses, processedRecords, statusSheet, mainSheet) {
  // 获取记录ID列的索引
  const headers = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
  const recordIdColumnIndex = headers.findIndex(header => header.trim() === '记录ID');
  
  if (recordIdColumnIndex === -1) {
    Logger.log(`警告：正式表中没有"记录ID"列`);
    return;
  }
  
  for (const course of courses) {
    let recordId = course.recordId;
    
    // 如果正式表中已有记录ID，使用它
    if (recordId) {
      continue;
    }
    
    // 尝试通过行号从状态表中获取记录ID
    const statusRow = statusSheet.getRange(course.rowIndex, 1, 1, 14).getValues()[0];
    if (statusRow[0]) {
      // 状态表中已有记录ID，使用它并更新正式表
      recordId = statusRow[0];
      course.recordId = recordId;
      mainSheet.getRange(course.rowIndex, recordIdColumnIndex + 1).setValue(recordId);
      Logger.log(`从状态表获取记录ID并更新正式表: ${recordId} (第${course.rowIndex}行)`);
      continue;
    }
    
    // 尝试通过key查找（向后兼容）
    const key = `${course.lessonNumber}_${course.date}`;
    const existingRecord = processedRecords.byKey.get(key);
    if (existingRecord && existingRecord.recordId) {
      recordId = existingRecord.recordId;
      course.recordId = recordId;
      mainSheet.getRange(course.rowIndex, recordIdColumnIndex + 1).setValue(recordId);
      Logger.log(`从状态表（通过key）获取记录ID并更新正式表: ${recordId} (第${course.rowIndex}行)`);
      continue;
    }
    
    // 生成新的记录ID
    recordId = generateRecordId();
    course.recordId = recordId;
    mainSheet.getRange(course.rowIndex, recordIdColumnIndex + 1).setValue(recordId);
    Logger.log(`为新记录生成ID并写入正式表: ${recordId} (第${course.rowIndex}行)`);
  }
}

/**
 * 生成唯一记录ID
 */
function generateRecordId() {
  // 使用时间戳 + 随机数生成唯一ID
  const timestamp = new Date().getTime();
  const random = Math.floor(Math.random() * 10000);
  return `REC_${timestamp}_${random}`;
}

/**
 * 获取已有的事件ID和token
 */
function getExistingEventIds(statusSheet, course) {
  // 优先通过记录ID查找
  let existingRecord = null;
  const processedRecords = readProcessedStatus(statusSheet);
  
  if (course.recordId) {
    existingRecord = processedRecords.byId.get(course.recordId);
  }
  
  // 如果没有找到，尝试通过key查找（向后兼容）
  if (!existingRecord) {
    const key = `${course.lessonNumber}_${course.date}`;
    existingRecord = processedRecords.byKey.get(key);
  }
  
  return {
    teacherEventId: existingRecord ? (existingRecord.teacherEventId || null) : null,
    studentEventId: existingRecord ? (existingRecord.studentEventId || null) : null,
    token: existingRecord ? (existingRecord.token || null) : null,
    hasChanges: existingRecord ? (existingRecord.token !== course.token) : true
  };
}

/**
 * 查找被删除的记录（状态表中有但正式表中没有的记录）
 * 通过记录ID匹配
 */
function findDeletedRecords(courses, processedRecords, statusSheet) {
  const deletedRecords = [];
  
  // 创建正式表中所有记录的ID集合
  const courseIds = new Set();
  courses.forEach(course => {
    if (course.recordId) {
      courseIds.add(course.recordId);
    }
  });
  
  // 检查状态表中的每条记录是否还在正式表中（通过记录ID匹配）
  processedRecords.byId.forEach((record, recordId) => {
    if (recordId && !courseIds.has(recordId)) {
      // 这条记录在状态表中但不在正式表中，说明被删除了
      deletedRecords.push({
        recordId: recordId,
        lessonNumber: record.lessonNumber,
        date: record.date,
        teacherCalendarId: record.teacherCalendarId || '',
        teacherEventId: record.teacherEventId || '',
        studentCalendarId: record.studentCalendarId || '',
        studentEventId: record.studentEventId || '',
        rowIndex: record.rowIndex,
        token: record.token || ''
      });
    }
  });
  
  // 检查通过key索引的记录（向后兼容，处理没有记录ID的旧记录）
  const courseKeys = new Set();
  courses.forEach(course => {
    const key = `${course.lessonNumber}_${course.date}`;
    courseKeys.add(key);
  });
  
  processedRecords.byKey.forEach((record, key) => {
    // 如果已经有记录ID且已处理过，跳过
    if (record.recordId && courseIds.has(record.recordId)) {
      return;
    }
    
    // 如果没有记录ID，通过key检查（向后兼容）
    if (!record.recordId && !courseKeys.has(key)) {
      deletedRecords.push({
        recordId: record.recordId || '',
        lessonNumber: record.lessonNumber,
        date: record.date,
        teacherCalendarId: record.teacherCalendarId || '',
        teacherEventId: record.teacherEventId || '',
        studentCalendarId: record.studentCalendarId || '',
        studentEventId: record.studentEventId || '',
        rowIndex: record.rowIndex,
        token: record.token || ''
      });
    }
  });
  
  return deletedRecords;
}

/**
 * 取消课程（删除日历事件并发送取消邮件）
 */
function cancelCourse(deletedRecord, statusSheet) {
  // 从状态表中获取日历ID和事件ID信息
  // deletedRecord 已经包含了 teacherEventId 和 studentEventId
  // 还需要获取日历ID（老师日历ID和学生日历ID）
  
  // 读取状态表中的完整信息（作为备用）
  const statusRow = statusSheet.getRange(deletedRecord.rowIndex, 1, 1, 16).getValues()[0];
  
  // 获取日历ID（优先使用deletedRecord中的，如果为空则从状态表中读取）
  const teacherCalendarId = deletedRecord.teacherCalendarId || statusRow[6] || ''; // 老师日历ID
  const studentCalendarId = deletedRecord.studentCalendarId || statusRow[11] || ''; // 学生日历ID
  
  // 1. 删除老师日历事件
  if (deletedRecord.teacherEventId) {
    try {
      if (teacherCalendarId) {
        // 如果有日历ID，直接删除
        deleteCalendarEvent(teacherCalendarId, deletedRecord.teacherEventId);
        Logger.log(`删除老师日历事件成功: ${deletedRecord.teacherEventId} (日历: ${teacherCalendarId})`);
      } else {
        // 如果没有日历ID，尝试通过事件ID删除（遍历所有日历）
        deleteCalendarEventById(deletedRecord.teacherEventId);
        Logger.log(`删除老师日历事件成功: ${deletedRecord.teacherEventId}`);
      }
    } catch (error) {
      Logger.log(`删除老师日历事件失败: ${deletedRecord.teacherEventId} - ${error.message}`);
    }
  }
  
  // 2. 删除学生日历事件
  if (deletedRecord.studentEventId) {
    try {
      if (studentCalendarId) {
        // 如果有日历ID，直接删除
        deleteCalendarEvent(studentCalendarId, deletedRecord.studentEventId);
        Logger.log(`删除学生日历事件成功: ${deletedRecord.studentEventId} (日历: ${studentCalendarId})`);
      } else {
        // 如果没有日历ID，尝试通过事件ID删除（遍历所有日历）
        deleteCalendarEventById(deletedRecord.studentEventId);
        Logger.log(`删除学生日历事件成功: ${deletedRecord.studentEventId}`);
      }
    } catch (error) {
      Logger.log(`删除学生日历事件失败: ${deletedRecord.studentEventId} - ${error.message}`);
    }
  }
  
  // 3. 发送取消邮件（需要从日历事件中获取参与者信息）
  // 由于记录已被删除，我们无法获取邮箱信息
  // 可以通过日历事件获取参与者信息
  try {
    sendCancellationEmails(deletedRecord);
  } catch (error) {
    Logger.log(`发送取消邮件失败: ${error.message}`);
  }
  
  // 4. 清空状态记录（保留行，但清空内容）
  const emptyRow = ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']; // 16列（包含记录ID和日历ID）
  statusSheet.getRange(deletedRecord.rowIndex, 1, 1, emptyRow.length).setValues([emptyRow]);
}

/**
 * 通过事件ID删除日历事件（尝试所有可能的日历）
 */
function deleteCalendarEventById(eventId) {
  if (!eventId) {
    return;
  }
  
  // 获取所有可访问的日历
  const calendars = CalendarApp.getAllCalendars();
  
  for (const calendar of calendars) {
    try {
      const event = calendar.getEventById(eventId);
      if (event) {
        event.deleteEvent();
        Logger.log(`删除日历事件成功: ${eventId} (日历: ${calendar.getName()})`);
        return; // 找到并删除后退出
      }
    } catch (error) {
      // 继续尝试下一个日历
      continue;
    }
  }
  
  Logger.log(`未找到日历事件: ${eventId}`);
}

/**
 * 发送课程取消邮件
 */
function sendCancellationEmails(deletedRecord) {
  // 由于记录已被删除，我们需要从日历事件中获取参与者信息
  // 或者从状态表中获取之前保存的信息
  
  // 尝试从日历事件中获取参与者信息
  const calendars = CalendarApp.getAllCalendars();
  let event = null;
  let calendar = null;
  
  // 先尝试通过老师日历事件ID获取
  if (deletedRecord.teacherEventId) {
    for (const cal of calendars) {
      try {
        event = cal.getEventById(deletedRecord.teacherEventId);
        if (event) {
          calendar = cal;
          break;
        }
      } catch (error) {
        continue;
      }
    }
  }
  
  // 如果没找到，尝试通过学生日历事件ID获取
  if (!event && deletedRecord.studentEventId) {
    for (const cal of calendars) {
      try {
        event = cal.getEventById(deletedRecord.studentEventId);
        if (event) {
          calendar = cal;
          break;
        }
      } catch (error) {
        continue;
      }
    }
  }
  
  if (!event) {
    Logger.log(`无法获取日历事件信息，跳过发送取消邮件`);
    return;
  }
  
  // 从事件中获取参与者信息
  const guests = event.getGuestList();
  const teacherEmail = guests.length > 0 ? guests[0].getEmail() : null;
  const studentEmail = guests.length > 1 ? guests[1].getEmail() : null;
  
  if (!teacherEmail && !studentEmail) {
    Logger.log(`无法获取参与者邮箱，跳过发送取消邮件`);
    return;
  }
  
  // 构建取消邮件内容
  const courseTitle = event.getTitle() || '课程';
  const eventDate = event.getStartTime();
  const dateStr = Utilities.formatDate(eventDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  // 发送给老师
  if (teacherEmail) {
    try {
      const subject = `课程取消通知：${courseTitle}`;
      const body = `
        <html>
          <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
            <h2 style="color: #d32f2f;">课程取消通知</h2>
            <p>您好，</p>
            <p>很遗憾地通知您，以下课程已被取消：</p>
            <div style="background-color: #f5f5f5; padding: 15px; border-radius: 5px; margin: 20px 0;">
              <p><strong>课程主题：</strong>${courseTitle}</p>
              <p><strong>原定日期：</strong>${dateStr}</p>
            </div>
            <p>课程事件已从您的日历中删除。</p>
            <p>如有任何问题，请及时联系。</p>
            <p style="margin-top: 30px; color: #666; font-size: 12px;">此邮件由系统自动发送，请勿回复。</p>
          </body>
        </html>
      `;
      
      MailApp.sendEmail({
        to: teacherEmail,
        subject: subject,
        htmlBody: body
      });
      
      Logger.log(`取消邮件发送成功（老师）: ${teacherEmail}`);
    } catch (error) {
      Logger.log(`取消邮件发送失败（老师）: ${teacherEmail} - ${error.message}`);
    }
  }
  
  // 发送给学生
  if (studentEmail) {
    try {
      const subject = `课程取消通知：${courseTitle}`;
      const body = `
        <html>
          <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
            <h2 style="color: #d32f2f;">课程取消通知</h2>
            <p>您好，</p>
            <p>很遗憾地通知您，以下课程已被取消：</p>
            <div style="background-color: #f5f5f5; padding: 15px; border-radius: 5px; margin: 20px 0;">
              <p><strong>课程主题：</strong>${courseTitle}</p>
              <p><strong>原定日期：</strong>${dateStr}</p>
            </div>
            <p>课程事件已从您的日历中删除。</p>
            <p>如有任何问题，请及时联系。</p>
            <p style="margin-top: 30px; color: #666; font-size: 12px;">此邮件由系统自动发送，请勿回复。</p>
          </body>
        </html>
      `;
      
      MailApp.sendEmail({
        to: studentEmail,
        subject: subject,
        htmlBody: body
      });
      
      Logger.log(`取消邮件发送成功（学生）: ${studentEmail}`);
    } catch (error) {
      Logger.log(`取消邮件发送失败（学生）: ${studentEmail} - ${error.message}`);
    }
  }
}

/**
 * 查找相同课次但不同日期的旧记录（用于检测日期变化）
 */
function findOldRecordsByLessonNumber(statusSheet, lessonNumber, currentDate) {
  const oldRecords = [];
  
  if (!statusSheet || statusSheet.getLastRow() < 2) {
    return oldRecords;
  }
  
  const dataRange = statusSheet.getDataRange();
  const values = dataRange.getValues();
  
  // 标准化当前日期用于比较
  const currentDateStr = currentDate instanceof Date ?
    Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') :
    String(currentDate);
  
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowLessonNumber = row[1]; // 课次在第1列（索引1）
    const rowDate = row[2]; // 日期在第2列（索引2）
    
    // 如果课次相同但日期不同
    if (rowLessonNumber === lessonNumber && rowDate) {
      const rowDateStr = rowDate instanceof Date ?
        Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') :
        String(rowDate);
      
      if (rowDateStr !== currentDateStr) {
        oldRecords.push({
          lessonNumber: rowLessonNumber,
          date: rowDate,
          teacherCalendarId: row[6] || '', // 老师日历ID在第6列
          teacherEventId: row[7] || '', // 老师日历事件ID在第7列（向后兼容：尝试旧格式）
          studentCalendarId: row[11] || '', // 学生日历ID在第11列
          studentEventId: row[12] || '', // 学生日历事件ID在第12列（向后兼容：尝试旧格式）
          rowIndex: i + 1
        });
        
        // 向后兼容：如果新格式没有事件ID，尝试从旧格式读取
        if (!oldRecords[oldRecords.length - 1].teacherEventId && row[5]) {
          oldRecords[oldRecords.length - 1].teacherEventId = row[5]; // 旧格式：老师日历事件ID在第5列
        }
        if (!oldRecords[oldRecords.length - 1].studentEventId && row[9]) {
          oldRecords[oldRecords.length - 1].studentEventId = row[9]; // 旧格式：学生日历事件ID在第9列
        }
      }
    }
  }
  
  return oldRecords;
}

/**
 * 删除旧状态记录
 */
function deleteOldStatusRecords(statusSheet, oldRecords) {
  // 从后往前删除，避免索引变化
  const rowsToDelete = oldRecords.map(r => r.rowIndex).sort((a, b) => b - a);
  
  for (const rowIndex of rowsToDelete) {
    try {
      statusSheet.deleteRow(rowIndex);
      Logger.log(`删除旧状态记录: 第${rowIndex}行`);
    } catch (error) {
      Logger.log(`删除旧状态记录失败: 第${rowIndex}行 - ${error.message}`);
    }
  }
}

/**
 * 获取日历（通过ID或邮箱，使用多种方法尝试）
 * 
 * 注意：CalendarApp.getCalendarById() 可能返回 null 而不是抛出异常
 * 如果日历ID是邮箱地址，可能需要特殊处理
 */
function getCalendarByIdOrEmail(calendarId, course) {
  if (!calendarId) {
    return null;
  }
  
  let calendar = null;
  
  // 方法1: 直接通过ID获取（这是最常用的方法）
  // 注意：即使 getAllCalendars() 不返回共享的日历，getCalendarById() 也可能可以访问
  try {
    calendar = CalendarApp.getCalendarById(calendarId);
    if (calendar) {
      Logger.log(`✓ 通过ID获取日历成功: ${calendarId} (${calendar.getName()})`);
      return calendar;
    } else {
      Logger.log(`✗ 通过ID获取日历返回null: ${calendarId}`);
    }
  } catch (error) {
    Logger.log(`✗ 通过ID获取日历抛出异常: ${calendarId} - ${error.message}`);
  }
  
  // 方法1.5: 尝试不同的ID格式（如果calendarId是邮箱）
  if (calendarId.includes('@')) {
    // 尝试添加 #gmail.com 后缀
    const idWithSuffix = calendarId + '#gmail.com';
    try {
      calendar = CalendarApp.getCalendarById(idWithSuffix);
      if (calendar) {
        Logger.log(`✓ 通过ID（带后缀）获取日历成功: ${idWithSuffix} (${calendar.getName()})`);
        return calendar;
      }
    } catch (error) {
      Logger.log(`✗ 通过ID（带后缀）获取日历失败: ${idWithSuffix} - ${error.message}`);
    }
    
    // 尝试使用邮箱作为ID（去掉可能的域名部分）
    const emailParts = calendarId.split('@');
    if (emailParts.length === 2) {
      const emailId = emailParts[0] + '@gmail.com';
      if (emailId !== calendarId) {
        try {
          calendar = CalendarApp.getCalendarById(emailId);
          if (calendar) {
            Logger.log(`✓ 通过邮箱ID获取日历成功: ${emailId} (${calendar.getName()})`);
            return calendar;
          }
        } catch (error) {
          Logger.log(`✗ 通过邮箱ID获取日历失败: ${emailId} - ${error.message}`);
        }
      }
    }
  }
  
  // 方法2: 从课程信息中获取对应的邮箱并尝试（如果calendarId不是邮箱）
  if (course && !calendarId.includes('@')) {
    // 如果calendarId不是邮箱，尝试从课程信息中获取邮箱
    const emailToTry = course.teacherCalendarId === calendarId ? 
                       course.teacherEmail : 
                       (course.studentCalendarId === calendarId ? course.studentEmail : null);
    
    if (emailToTry) {
      try {
        calendar = CalendarApp.getCalendarById(emailToTry);
        if (calendar) {
          Logger.log(`✓ 通过课程邮箱获取日历成功: ${emailToTry} (${calendar.getName()})`);
          return calendar;
        }
      } catch (error) {
        Logger.log(`✗ 通过课程邮箱获取日历失败: ${emailToTry} - ${error.message}`);
      }
    }
  }
  
  Logger.log(`✗ 无法找到日历: ${calendarId}，请检查：1) 日历ID是否正确 2) 是否有访问权限 3) 日历是否已共享`);
  return null;
}

/**
 * 删除日历事件（通过日历ID和事件ID）
 */
function deleteCalendarEvent(calendarId, eventId) {
  if (!calendarId || !eventId) {
    return;
  }
  
  try {
    // 使用更健壮的获取日历方法
    const calendar = getCalendarByIdOrEmail(calendarId, null);
    if (!calendar) {
      Logger.log(`找不到日历: ${calendarId}`);
      return;
    }
    
    const event = calendar.getEventById(eventId);
    if (event) {
      event.deleteEvent();
      Logger.log(`删除日历事件成功: ${eventId} (日历: ${calendarId})`);
    } else {
      Logger.log(`找不到日历事件: ${eventId} (日历: ${calendarId})`);
    }
  } catch (error) {
    Logger.log(`删除日历事件失败: ${eventId} (日历: ${calendarId}) - ${error.message}`);
  }
}

/**
 * 计算课程关键信息的token（用于检测变化）
 * 包括：日期、开始时间、结束时间、课程内容、老师、老师邮箱、学生、学生邮箱
 */
function calculateCourseToken(course) {
  // 标准化日期和时间格式
  const dateStr = course.date instanceof Date ? 
    Utilities.formatDate(course.date, Session.getScriptTimeZone(), 'yyyy-MM-dd') : 
    String(course.date);
  
  const startTimeStr = course.startTime instanceof Date ?
    Utilities.formatDate(course.startTime, Session.getScriptTimeZone(), 'HH:mm') :
    String(course.startTime);
  
  const endTimeStr = course.endTime instanceof Date ?
    Utilities.formatDate(course.endTime, Session.getScriptTimeZone(), 'HH:mm') :
    String(course.endTime);
  
  // 构建关键信息字符串
  const keyInfo = [
    dateStr,
    startTimeStr,
    endTimeStr,
    String(course.courseTitle || ''),
    String(course.teacherName || ''),
    String(course.teacherEmail || ''),
    String(course.studentName || ''),
    String(course.studentEmail || '')
  ].join('|');
  
  // 计算MD5哈希作为token
  const hash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    keyInfo,
    Utilities.Charset.UTF_8
  );
  
  // 转换为十六进制字符串
  const token = hash.map(function(byte) {
    return ('0' + (byte & 0xFF).toString(16)).slice(-2);
  }).join('');
  
  return token;
}

// ==================== 邮件发送模块 ====================

/**
 * 发送课程邮件
 */
function sendCourseEmail(recipientEmail, recipientName, course, otherPartyName) {
  if (!recipientEmail) {
    throw new Error('收件人邮箱为空');
  }
  
  const subject = CONFIG.EMAIL_TEMPLATE.subject.replace('{courseTitle}', course.courseTitle);
  
  const body = CONFIG.EMAIL_TEMPLATE.body
    .replace(/{recipientName}/g, recipientName)
    .replace(/{courseTitle}/g, course.courseTitle)
    .replace(/{courseDate}/g, formatDate(course.date))
    .replace(/{startTime}/g, course.startTime)
    .replace(/{endTime}/g, course.endTime)
    .replace(/{teacherName}/g, course.teacherName)
    .replace(/{studentName}/g, course.studentName);
  
  MailApp.sendEmail({
    to: recipientEmail,
    subject: subject,
    htmlBody: body
  });
}

// ==================== 日历事件创建模块 ====================

/**
 * 创建或更新日历事件
 * @param {string} calendarId - 日历ID
 * @param {Object} course - 课程对象
 * @param {string|null} existingEventId - 已有的事件ID（如果存在则更新，否则创建）
 * @returns {string} 事件ID
 */
function createOrUpdateCalendarEvent(calendarId, course, existingEventId) {
  if (!calendarId) {
    throw new Error('日历ID为空');
  }
  
  // 解析日期和时间
  const startDateTime = parseDateTime(course.date, course.startTime);
  const endDateTime = parseDateTime(course.date, course.endTime);
  
  if (!startDateTime || !endDateTime) {
    throw new Error('日期时间解析失败');
  }
  
  // 获取日历（直接通过ID获取，不遍历，不使用默认日历）
  const calendar = getCalendarByIdOrEmail(calendarId, course);
  
  if (!calendar) {
    throw new Error(`找不到日历: ${calendarId}，请检查：1) 日历ID是否正确 2) 是否有访问权限 3) 日历是否已共享`);
  }
  
  // 记录实际使用的日历信息
  Logger.log(`使用日历: ${calendar.getName()} (${calendar.getId()})，目标ID: ${calendarId}`);
  
  // 构建事件信息
  const eventSummary = course.courseTitle;
  const eventDescription = `课程：${course.courseTitle}\n老师：${course.teacherName}\n学生：${course.studentName}\n课次：${course.lessonNumber}`;
  const eventStart = new Date(startDateTime);
  const eventEnd = new Date(endDateTime);
  const eventGuests = `${course.teacherEmail},${course.studentEmail}`;
  
  let event;
  
  if (existingEventId) {
    // 更新已有事件
    try {
      event = calendar.getEventById(existingEventId);
      
      // 更新事件信息
      event.setTitle(eventSummary);
      event.setDescription(eventDescription);
      event.setTime(eventStart, eventEnd);
      
      // 更新参与者（使用正确的方法）
      // 先获取现有参与者列表
      const existingGuests = event.getGuestList();
      const existingEmails = existingGuests.map(guest => guest.getEmail());
      const newEmails = eventGuests.split(',').map(email => email.trim());
      
      // 添加新参与者
      for (const email of newEmails) {
        if (email && !existingEmails.includes(email)) {
          event.addGuest(email);
        }
      }
      
      // 移除不在新列表中的参与者（可选，根据需求决定）
      // 这里不删除，只添加新的参与者
      
      Logger.log(`更新日历事件: ${existingEventId}`);
      return existingEventId;
    } catch (error) {
      // 如果事件不存在或无法访问，则创建新事件
      Logger.log(`无法更新事件 ${existingEventId}，将创建新事件: ${error.message}`);
      // 继续执行创建逻辑
    }
  }
  
  // 创建新事件
  event = calendar.createEvent(
    eventSummary,
    eventStart,
    eventEnd,
    {
      description: eventDescription,
      guests: eventGuests,
      sendInvites: true
    }
  );
  
  Logger.log(`创建新日历事件: ${event.getId()}`);
  return event.getId();
}

/**
 * 创建日历事件（保留向后兼容）
 * @deprecated 使用 createOrUpdateCalendarEvent 代替
 */
function createCalendarEvent(calendarId, course) {
  return createOrUpdateCalendarEvent(calendarId, course, null);
}

// ==================== 状态记录模块 ====================

/**
 * 确保状态表存在
 */
function ensureStatusSheet(spreadsheet) {
  let statusSheet = spreadsheet.getSheetByName(CONFIG.STATUS_SHEET_NAME);
  
  if (!statusSheet) {
    // 创建隐藏表
    statusSheet = spreadsheet.insertSheet(CONFIG.STATUS_SHEET_NAME);
    statusSheet.hideSheet(); // 隐藏表
    
    // 设置表头（索引表结构）
    const headers = [
      '记录ID',            // 0 - 唯一标识符（用于正式表和索引表一一对应）
      '课次',              // 1 - 索引字段
      '日期',              // 2 - 索引字段
      'Token',             // 3 - 关键信息哈希值（用于检测变化）
      '老师邮件状态',      // 4
      '老师邮件发送时间',  // 5
      '老师日历ID',        // 6 - 老师日历ID（用于删除事件）
      '老师日历事件ID',    // 7 - 老师日历事件ID
      '老师日历创建时间',  // 8
      '学生邮件状态',      // 9
      '学生邮件发送时间',  // 10
      '学生日历ID',        // 11 - 学生日历ID（用于删除事件）
      '学生日历事件ID',    // 12 - 学生日历事件ID
      '学生日历创建时间',  // 13
      '处理状态',          // 14
      '最后更新时间'       // 15
    ];
    
    statusSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    statusSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4285F4')
      .setFontColor('#FFFFFF');
    
    // 冻结首行
    statusSheet.setFrozenRows(1);
    
    Logger.log(`创建状态表: ${CONFIG.STATUS_SHEET_NAME}`);
  }
  
  return statusSheet;
}

/**
 * 同步状态表，确保和正式表一一对应
 * 状态表的第i行对应正式表的第i+1行（正式表有表头）
 */
function syncStatusSheet(statusSheet, courseCount) {
  const currentRowCount = statusSheet.getLastRow();
  const targetRowCount = courseCount + 1; // +1 是表头行
  
  if (currentRowCount < targetRowCount) {
    // 需要添加行
    const rowsToAdd = targetRowCount - currentRowCount;
    const emptyRow = ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']; // 16列（包含记录ID和日历ID）
    const rows = [];
    for (let i = 0; i < rowsToAdd; i++) {
      rows.push(emptyRow);
    }
    statusSheet.getRange(currentRowCount + 1, 1, rowsToAdd, emptyRow.length).setValues(rows);
    Logger.log(`状态表同步：添加了 ${rowsToAdd} 行`);
  } else if (currentRowCount > targetRowCount) {
    // 需要删除多余的行（保留表头）
    const rowsToDelete = currentRowCount - targetRowCount;
    statusSheet.deleteRows(targetRowCount + 1, rowsToDelete);
    Logger.log(`状态表同步：删除了 ${rowsToDelete} 行`);
  }
}

/**
 * 更新状态记录（通过行号索引，和正式表一一对应）
 */
function updateStatusRecord(statusSheet, course, result) {
  const now = new Date();
  const nowStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  
  // 使用course.rowIndex来确定状态表的行号
  // 状态表的第i行对应正式表的第i+1行（正式表有表头，状态表也有表头）
  const rowIndex = course.rowIndex; // course.rowIndex是正式表的行号（从1开始，包含表头）
  
  // 读取当前行的现有记录（如果有）
  let existingRecord = null;
  if (rowIndex <= statusSheet.getLastRow()) {
    const rowValues = statusSheet.getRange(rowIndex, 1, 1, 16).getValues()[0];
    if (rowValues[1] || rowValues[2]) { // 如果课次或日期不为空，说明有记录
      existingRecord = rowValues;
    }
  }
  
  // 获取或生成记录ID
  const recordId = course.recordId || (existingRecord ? (existingRecord[0] || generateRecordId()) : generateRecordId());
  
  // 保留已有的事件ID和日历ID（如果更新失败）
  const teacherCalendarId = course.teacherCalendarId || (existingRecord ? (existingRecord[6] || '') : '');
  const teacherEventId = result.teacherEmail.eventId || (existingRecord ? (existingRecord[7] || '') : '');
  const studentCalendarId = course.studentCalendarId || (existingRecord ? (existingRecord[11] || '') : '');
  const studentEventId = result.studentEmail.eventId || (existingRecord ? (existingRecord[12] || '') : '');
  
  // 如果事件ID存在，更新创建时间；如果是新创建的，使用当前时间；如果是已有的，保留原时间
  let teacherEventTime = '';
  let studentEventTime = '';
  
  if (result.teacherEmail.eventId) {
    // 新创建或更新的事件
    teacherEventTime = nowStr;
  } else if (existingRecord && existingRecord[7]) {
    // 保留原有的创建时间（向后兼容：尝试旧格式）
    teacherEventTime = existingRecord[8] || existingRecord[6] || '';
  }
  
  if (result.studentEmail.eventId) {
    // 新创建或更新的事件
    studentEventTime = nowStr;
  } else if (existingRecord && existingRecord[12]) {
    // 保留原有的创建时间（向后兼容：尝试旧格式）
    studentEventTime = existingRecord[13] || existingRecord[10] || '';
  }
  
  // 获取或计算token
  const token = course.token || calculateCourseToken(course);
  
  const record = [
    recordId, // 0 - 记录ID（唯一标识符）
    course.lessonNumber, // 1 - 课次（索引字段）
    course.date, // 2 - 日期（索引字段）
    token, // 3 - Token（关键信息哈希值）
    result.teacherEmail.sent ? '已发送' : (result.teacherEmail.error || (existingRecord ? existingRecord[4] : '未发送')), // 4 - 老师邮件状态
    result.teacherEmail.sent ? nowStr : (existingRecord ? existingRecord[5] : ''), // 5 - 老师邮件发送时间
    teacherCalendarId, // 6 - 老师日历ID（用于删除事件）
    teacherEventId, // 7 - 老师日历事件ID（保留已有的）
    teacherEventTime, // 8 - 老师日历创建/更新时间
    result.studentEmail.sent ? '已发送' : (result.studentEmail.error || (existingRecord ? existingRecord[9] : '未发送')), // 9 - 学生邮件状态
    result.studentEmail.sent ? nowStr : (existingRecord ? existingRecord[10] : ''), // 10 - 学生邮件发送时间
    studentCalendarId, // 11 - 学生日历ID（用于删除事件）
    studentEventId, // 12 - 学生日历事件ID（保留已有的）
    studentEventTime, // 13 - 学生日历创建/更新时间
    result.status, // 14 - 处理状态
    nowStr // 15 - 最后更新时间
  ];
  
  // 直接更新对应行（状态表和正式表一一对应）
  statusSheet.getRange(rowIndex, 1, 1, record.length).setValues([record]);
}

// ==================== 工具函数 ====================

/**
 * 解析日期时间
 */
function parseDateTime(dateInput, timeInput) {
  try {
    let date;
    let hours = 0;
    let minutes = 0;
    let seconds = 0;
    
    // 处理日期：可能是 Date 对象或字符串
    if (dateInput instanceof Date) {
      // 如果是 Date 对象，直接使用
      date = new Date(dateInput);
    } else if (typeof dateInput === 'string') {
      // 解析日期字符串：支持 2025/11/13 或 2025-11-13 格式
      if (dateInput.includes('/')) {
        const [year, month, day] = dateInput.split('/').map(Number);
        date = new Date(year, month - 1, day);
      } else if (dateInput.includes('-')) {
        date = new Date(dateInput);
      } else {
        throw new Error(`不支持的日期格式: ${dateInput}`);
      }
    } else {
      throw new Error(`不支持的日期类型: ${typeof dateInput}`);
    }
    
    // 处理时间：可能是 Date 对象或字符串
    if (timeInput instanceof Date) {
      // Google Sheets 时间列返回的 Date 对象（通常是 1899-12-30 + 时间）
      hours = timeInput.getHours();
      minutes = timeInput.getMinutes();
      seconds = timeInput.getSeconds();
    } else if (typeof timeInput === 'string') {
      // 解析时间字符串：支持 10:00 或 10:00:00 格式
      const timeParts = timeInput.split(':').map(Number);
      hours = timeParts[0] || 0;
      minutes = timeParts[1] || 0;
      seconds = timeParts[2] || 0;
    } else if (typeof timeInput === 'number') {
      // 可能是小数形式的时间（0-1之间，表示一天中的时间）
      const totalSeconds = Math.round(timeInput * 24 * 60 * 60);
      hours = Math.floor(totalSeconds / 3600);
      minutes = Math.floor((totalSeconds % 3600) / 60);
      seconds = totalSeconds % 60;
    } else {
      throw new Error(`不支持的时间类型: ${typeof timeInput}`);
    }
    
    // 设置时间
    date.setHours(hours, minutes, seconds, 0);
    
    return date;
  } catch (error) {
    Logger.log(`日期时间解析错误: ${dateInput} (${typeof dateInput}) ${timeInput} (${typeof timeInput}) - ${error.message}`);
    return null;
  }
}

/**
 * 格式化日期显示
 */
function formatDate(dateInput) {
  try {
    // 如果是 Date 对象，格式化为字符串
    if (dateInput instanceof Date) {
      const year = dateInput.getFullYear();
      const month = String(dateInput.getMonth() + 1).padStart(2, '0');
      const day = String(dateInput.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    }
    
    // 如果是字符串
    if (typeof dateInput === 'string') {
      if (dateInput.includes('/')) {
        return dateInput.replace(/\//g, '-');
      }
      return dateInput;
    }
    
    return String(dateInput);
  } catch (error) {
    return String(dateInput);
  }
}

// ==================== 测试函数 ====================

/**
 * 测试函数 - 处理单条记录
 */
function testSingleRecord() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  ensureStatusSheet(spreadsheet);
  
  const mainSheet = spreadsheet.getSheetByName(CONFIG.MAIN_SHEET_NAME);
  const courses = readCourseData(mainSheet);
  
  if (courses.length > 0) {
    const statusSheet = spreadsheet.getSheetByName(CONFIG.STATUS_SHEET_NAME);
    const result = processCourse(courses[0], statusSheet);
    Logger.log(JSON.stringify(result, null, 2));
  } else {
    Logger.log('没有找到课程数据');
  }
}

/**
 * 测试函数 - 读取数据
 */
function testReadData() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = spreadsheet.getSheetByName(CONFIG.MAIN_SHEET_NAME);
  const courses = readCourseData(mainSheet);
  Logger.log(`读取到 ${courses.length} 条记录`);
  Logger.log(JSON.stringify(courses, null, 2));
}


