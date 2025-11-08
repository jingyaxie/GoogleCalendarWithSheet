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
  // 主表名称（根据实际情况修改，向后兼容使用）
  MAIN_SHEET_NAME: '课程安排',
  
  // 配置表名称（用于管理要处理的 sheet 列表）
  CONFIG_SHEET_NAME: '_SheetConfig',
  
  // 隐藏状态表名称前缀（实际状态表名称 = STATUS_SHEET_PREFIX + Sheet名称）
  STATUS_SHEET_PREFIX: '_StatusLog_',
  
  // 时区设置
  TIMEZONE: 'Asia/Shanghai',
  
  // 速率限制配置
  RATE_LIMIT: {
    // 每次操作之间的延迟（毫秒）
    DELAY_BETWEEN_OPERATIONS: 500,
    // 重试次数
    MAX_RETRIES: 3,
    // 重试延迟（毫秒）
    RETRY_DELAY: 2000,
    // 速率限制错误的关键词
    RATE_LIMIT_KEYWORDS: ['too many', 'rate limit', 'quota', 'try again later']
  },
  
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

// ==================== 菜单功能 ====================

/**
 * 当打开表格时自动创建自定义菜单
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // 创建自定义菜单
  ui.createMenu('📅 课程同步')
    .addItem('🔄 执行同步', 'menuRunSync')
    .addSeparator()
    .addItem('📋 查看配置', 'menuViewConfig')
    .addItem('📊 查看状态表', 'menuViewStatus')
    .addSeparator()
    .addItem('ℹ️ 关于', 'menuAbout')
    .addToUi();
}

/**
 * 菜单项：执行同步
 */
function menuRunSync() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '确认执行同步',
      '这将处理所有配置的课程表，发送邮件并创建日历事件。\n\n是否继续？',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      // 执行主函数
      main();
      
      // 显示完成提示
      ui.alert(
        '同步完成',
        '课程同步已完成，请查看执行日志了解详细信息。',
        ui.ButtonSet.OK
      );
    }
  } catch (error) {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      '执行错误',
      '同步过程中发生错误：\n' + error.message,
      ui.ButtonSet.OK
    );
    Logger.log('菜单执行同步错误: ' + error.message);
  }
}

/**
 * 菜单项：查看配置
 */
function menuViewConfig() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = spreadsheet.getSheetByName(CONFIG.CONFIG_SHEET_NAME);
    
    if (!configSheet) {
      const ui = SpreadsheetApp.getUi();
      ui.alert(
        '配置表不存在',
        `找不到配置表 "${CONFIG.CONFIG_SHEET_NAME}"，请先创建配置表。`,
        ui.ButtonSet.OK
      );
      return;
    }
    
    // 激活配置表
    configSheet.activate();
    
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      '配置表已打开',
      '配置表已激活，请查看配置信息。',
      ui.ButtonSet.OK
    );
  } catch (error) {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      '查看配置错误',
      '查看配置时发生错误：\n' + error.message,
      ui.ButtonSet.OK
    );
    Logger.log('查看配置错误: ' + error.message);
  }
}

/**
 * 菜单项：查看状态表
 */
function menuViewStatus() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    
    // 读取配置表，获取所有启用的 Sheet
    const sheetConfigMap = readSheetConfig(spreadsheet);
    
    if (sheetConfigMap.size === 0) {
      ui.alert(
        '没有配置的 Sheet',
        '配置表中没有启用的 Sheet，请先配置。',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // 如果有多个 Sheet，让用户选择
    const sheetNames = Array.from(sheetConfigMap.keys());
    let selectedSheet = null;
    
    if (sheetNames.length === 1) {
      selectedSheet = sheetNames[0];
    } else {
      // 创建选择对话框
      const html = HtmlService.createHtmlOutput(`
        <div style="font-family: Arial, sans-serif; padding: 20px;">
          <h3>选择要查看的 Sheet</h3>
          <select id="sheetSelect" style="width: 100%; padding: 8px; margin: 10px 0;">
            ${sheetNames.map(name => `<option value="${name}">${name}</option>`).join('')}
          </select>
          <button onclick="google.script.host.close(); google.script.run('menuViewStatusSheet', document.getElementById('sheetSelect').value)" 
                  style="width: 100%; padding: 10px; background: #4285F4; color: white; border: none; border-radius: 4px; cursor: pointer;">
            查看状态表
          </button>
        </div>
      `)
        .setWidth(300)
        .setHeight(150);
      
      ui.showModalDialog(html, '选择 Sheet');
      return;
    }
    
    // 显示状态表
    menuViewStatusSheet(selectedSheet);
    
  } catch (error) {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      '查看状态表错误',
      '查看状态表时发生错误：\n' + error.message,
      ui.ButtonSet.OK
    );
    Logger.log('查看状态表错误: ' + error.message);
  }
}

/**
 * 查看指定 Sheet 的状态表
 */
function menuViewStatusSheet(sheetName) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const statusSheetName = CONFIG.STATUS_SHEET_PREFIX + sheetName;
    const statusSheet = spreadsheet.getSheetByName(statusSheetName);
    
    if (!statusSheet) {
      const ui = SpreadsheetApp.getUi();
      ui.alert(
        '状态表不存在',
        `找不到状态表 "${statusSheetName}"，请先执行一次同步。`,
        ui.ButtonSet.OK
      );
      return;
    }
    
    // 显示状态表（取消隐藏）
    statusSheet.showSheet();
    statusSheet.activate();
    
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      '状态表已打开',
      `状态表 "${statusSheetName}" 已激活并显示。`,
      ui.ButtonSet.OK
    );
  } catch (error) {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      '查看状态表错误',
      '查看状态表时发生错误：\n' + error.message,
      ui.ButtonSet.OK
    );
    Logger.log('查看状态表错误: ' + error.message);
  }
}

/**
 * 菜单项：关于
 */
function menuAbout() {
  const ui = SpreadsheetApp.getUi();
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 20px; line-height: 1.6;">
      <h2 style="color: #4285F4;">📅 课程同步系统</h2>
      <p><strong>版本：</strong>2.0</p>
      <p><strong>功能：</strong></p>
      <ul>
        <li>从配置表读取多个课程表</li>
        <li>自动发送邮件通知给老师和学生</li>
        <li>创建日历事件到老师和学生的日历</li>
        <li>跟踪处理状态和记录ID</li>
        <li>支持课程更新和删除</li>
      </ul>
      <p><strong>配置表：</strong>${CONFIG.CONFIG_SHEET_NAME}</p>
      <p><strong>状态表前缀：</strong>${CONFIG.STATUS_SHEET_PREFIX}</p>
      <hr>
      <p style="color: #666; font-size: 12px;">使用菜单中的"执行同步"来开始处理课程数据。</p>
    </div>
  `)
    .setWidth(400)
    .setHeight(400);
  
  ui.showModalDialog(html, '关于');
}

// ==================== 主函数 ====================

/**
 * 主执行函数 - 处理所有课程记录
 * 从配置表 _SheetConfig 读取要处理的 sheet 列表，然后循环处理每个 sheet
 */
function main() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // 从配置表读取要处理的 sheet 配置信息
    const sheetConfigMap = readSheetConfig(spreadsheet);
    
    if (sheetConfigMap.size === 0) {
      Logger.log('警告：没有找到需要处理的 sheet，请检查配置表 _SheetConfig');
      return;
    }
    
    Logger.log(`从配置表读取到 ${sheetConfigMap.size} 个需要处理的 sheet: ${Array.from(sheetConfigMap.keys()).join(', ')}`);
    
    // 循环处理每个 sheet
    const allResults = [];
    for (const [sheetName, config] of sheetConfigMap) {
      try {
        Logger.log(`\n========== 开始处理 Sheet: ${sheetName} ==========`);
        const result = processSheet(spreadsheet, sheetName, config);
        allResults.push({
          sheetName: sheetName,
          success: result.success,
          total: result.total,
          processed: result.processed,
          failed: result.failed,
          error: result.error
        });
        Logger.log(`========== Sheet ${sheetName} 处理完成 ==========\n`);
      } catch (error) {
        Logger.log(`处理 Sheet ${sheetName} 时发生错误: ${error.message}`);
        allResults.push({
          sheetName: sheetName,
          success: false,
          total: 0,
          processed: 0,
          failed: 0,
          error: error.message
        });
      }
    }
    
    // 输出汇总结果
    Logger.log('\n=== 所有 Sheet 处理结果汇总 ===');
    let totalSuccess = 0;
    let totalFailed = 0;
    let totalProcessed = 0;
    for (const result of allResults) {
      if (result.success) {
        totalSuccess++;
      } else {
        totalFailed++;
      }
      totalProcessed += result.processed;
      Logger.log(`${result.sheetName}: ${result.success ? '成功' : '失败'} - 处理 ${result.processed} 条记录${result.error ? ` (错误: ${result.error})` : ''}`);
    }
    Logger.log(`总计: 成功 ${totalSuccess}, 失败 ${totalFailed}, 共处理 ${totalProcessed} 条记录`);
    
  } catch (error) {
    Logger.log(`主函数执行失败: ${error.message}`);
    throw error;
  }
}

/**
 * 处理单个 Sheet 的所有课程记录
 * @param {Spreadsheet} spreadsheet - 表格对象
 * @param {string} sheetName - Sheet 名称
 * @param {Object} config - Sheet 配置信息（包含邮箱和日历ID）
 * @returns {Object} 处理结果
 */
function processSheet(spreadsheet, sheetName, config) {
  try {
    // 获取主表
    const mainSheet = spreadsheet.getSheetByName(sheetName);
    if (!mainSheet) {
      throw new Error(`找不到 Sheet: ${sheetName}`);
    }
    
    // 生成状态表名称
    const statusSheetName = CONFIG.STATUS_SHEET_PREFIX + sheetName;
    
    // 确保隐藏状态表存在
    ensureStatusSheet(spreadsheet, statusSheetName);
    
    // 确保正式表有"记录ID"列
    ensureRecordIdColumn(mainSheet);
    
    // 读取课程数据，传入配置信息（包含时区）
    const courses = readCourseData(mainSheet, config);
    // 为每条课程记录添加时区和提醒时间信息
    courses.forEach(course => {
      course.timezone = config.timezone;
      course.reminderMinutes = config.reminderMinutes;
    });
    Logger.log(`[${sheetName}] 读取到 ${courses.length} 条课程记录，时区: ${config.timezone}, 提醒时间: ${config.reminderMinutes ? config.reminderMinutes + '分钟' : '未配置'}`);
    
    // 读取已处理状态（在同步之前读取，以便检测被删除的记录）
    const statusSheet = spreadsheet.getSheetByName(statusSheetName);
    const processedRecords = readProcessedStatus(statusSheet);
    
    // 检测被删除的记录（在同步状态表之前检测，避免状态表被删除后无法检测）
    const deletedRecords = findDeletedRecords(courses, processedRecords, statusSheet);
    if (deletedRecords.length > 0) {
      Logger.log(`[${sheetName}] 检测到 ${deletedRecords.length} 条被删除的记录，将取消课程`);
      for (const deletedRecord of deletedRecords) {
        try {
          cancelCourse(deletedRecord, statusSheet);
          Logger.log(`[${sheetName}] 取消课程成功: ${deletedRecord.lessonNumber} - ${deletedRecord.date}`);
        } catch (error) {
          Logger.log(`[${sheetName}] 取消课程失败: ${deletedRecord.lessonNumber} - ${error.message}`);
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
        const timezone = course.timezone || CONFIG.TIMEZONE || Session.getScriptTimeZone();
        const oldRecords = findOldRecordsByLessonNumber(statusSheet, course.lessonNumber, course.date, timezone);
        if (oldRecords.length > 0) {
          Logger.log(`[${sheetName}] 检测到日期变化: ${course.lessonNumber}，将在处理时删除旧日期的日历事件`);
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
        Logger.log(`[${sheetName}] 检测到关键信息变化: ${course.lessonNumber} (旧token: ${existingToken}, 新token: ${currentToken})`);
        return true;
      }
      
      // token相同，说明关键信息没有变化
      // 检查是否已有日历事件ID，如果有则验证事件是否真实存在
      // 注意：只有当事件ID非空字符串时才检查
      const hasTeacherEventId = existingRecord.teacherEventId && String(existingRecord.teacherEventId).trim() !== '';
      const hasStudentEventId = existingRecord.studentEventId && String(existingRecord.studentEventId).trim() !== '';
      
      if (hasTeacherEventId || hasStudentEventId) {
        // 验证事件是否真实存在于日历中
        let teacherEventExists = false;
        let studentEventExists = false;
        let needRecreate = false;
        
        // 验证老师日历事件（只有当事件ID非空时才验证）
        if (hasTeacherEventId && existingRecord.teacherCalendarId) {
          try {
            teacherEventExists = verifyCalendarEventExists(existingRecord.teacherCalendarId, existingRecord.teacherEventId);
            if (!teacherEventExists) {
              Logger.log(`[${sheetName}] 老师日历事件不存在（可能被删除）: ${existingRecord.teacherEventId}，将重新创建`);
              needRecreate = true;
              // 更新状态表，清除无效的事件ID
              statusSheet.getRange(existingRecord.rowIndex, 8).setValue(''); // 第8列是老师日历事件ID
              existingRecord.teacherEventId = '';
            }
          } catch (error) {
            Logger.log(`[${sheetName}] 验证老师日历事件失败: ${existingRecord.teacherEventId} - ${error.message}`);
            teacherEventExists = false; // 验证失败，认为不存在
            needRecreate = true;
            // 更新状态表，清除无效的事件ID
            statusSheet.getRange(existingRecord.rowIndex, 8).setValue('');
            existingRecord.teacherEventId = '';
          }
        } else if (hasTeacherEventId) {
          // 有事件ID但没有日历ID，无法验证，需要重新创建
          Logger.log(`[${sheetName}] 老师日历事件ID存在但缺少日历ID，将重新创建`);
          needRecreate = true;
          statusSheet.getRange(existingRecord.rowIndex, 8).setValue('');
          existingRecord.teacherEventId = '';
        }
        
        // 验证学生日历事件（只有当事件ID非空时才验证）
        if (hasStudentEventId && existingRecord.studentCalendarId) {
          try {
            studentEventExists = verifyCalendarEventExists(existingRecord.studentCalendarId, existingRecord.studentEventId);
            if (!studentEventExists) {
              Logger.log(`[${sheetName}] 学生日历事件不存在（可能被删除）: ${existingRecord.studentEventId}，将重新创建`);
              needRecreate = true;
              // 更新状态表，清除无效的事件ID
              statusSheet.getRange(existingRecord.rowIndex, 13).setValue(''); // 第13列是学生日历事件ID
              existingRecord.studentEventId = '';
            }
          } catch (error) {
            Logger.log(`[${sheetName}] 验证学生日历事件失败: ${existingRecord.studentEventId} - ${error.message}`);
            studentEventExists = false; // 验证失败，认为不存在
            needRecreate = true;
            // 更新状态表，清除无效的事件ID
            statusSheet.getRange(existingRecord.rowIndex, 13).setValue('');
            existingRecord.studentEventId = '';
          }
        } else if (hasStudentEventId) {
          // 有事件ID但没有日历ID，无法验证，需要重新创建
          Logger.log(`[${sheetName}] 学生日历事件ID存在但缺少日历ID，将重新创建`);
          needRecreate = true;
          statusSheet.getRange(existingRecord.rowIndex, 13).setValue('');
          existingRecord.studentEventId = '';
        }
        
        // 如果两个事件都存在，才跳过处理
        // 注意：如果只有部分事件ID，也需要处理（创建缺失的事件）
        const hasValidTeacherEvent = hasTeacherEventId && existingRecord.teacherCalendarId && teacherEventExists;
        const hasValidStudentEvent = hasStudentEventId && existingRecord.studentCalendarId && studentEventExists;
        
        if (hasValidTeacherEvent && hasValidStudentEvent) {
          Logger.log(`[${sheetName}] 跳过处理（token相同且日历事件已验证存在）: ${course.lessonNumber}`);
          return false;
        }
        
        // 如果有事件不存在或需要重新创建，需要重新处理
        if (needRecreate || !teacherEventExists || !studentEventExists) {
          Logger.log(`[${sheetName}] 需要重新处理（日历事件不存在或需要创建）: ${course.lessonNumber}`);
          return true;
        }
      }
      
      // token相同但没有日历事件ID，可能是之前创建失败，需要重试
      // 但只有在状态不是已完成时才处理
      if (existingRecord.status !== '已完成') {
        Logger.log(`[${sheetName}] 重试处理（token相同但之前失败）: ${course.lessonNumber}`);
        return true;
      }
      
      // token相同且已完成，跳过
      return false;
    });
    
    Logger.log(`[${sheetName}] 需要处理 ${toProcess.length} 条记录`);
    
    // 处理每条记录
    const results = [];
    for (const course of toProcess) {
      try {
        const result = processCourse(course, statusSheet);
        results.push(result);
        Logger.log(`[${sheetName}] 处理完成: ${course.lessonNumber} - ${result.status}`);
      } catch (error) {
        Logger.log(`[${sheetName}] 处理失败: ${course.lessonNumber} - ${error.message}`);
        results.push({
          course: course,
          status: '失败',
          error: error.message
        });
      }
    }
    
    // 输出处理结果
    Logger.log(`\n[${sheetName}] === 处理结果汇总 ===`);
    let successCount = 0;
    let failedCount = 0;
    for (const result of results) {
      if (result.status === '已完成') {
        successCount++;
      } else {
        failedCount++;
      }
    }
    Logger.log(`[${sheetName}] 成功: ${successCount}, 失败: ${failedCount}`);
    
    return {
      success: true,
      total: courses.length,
      processed: toProcess.length,
      failed: failedCount
    };
    
  } catch (error) {
    Logger.log(`处理 Sheet ${sheetName} 失败: ${error.message}`);
    return {
      success: false,
      total: 0,
      processed: 0,
      failed: 0,
      error: error.message
    };
  }
}

/**
 * 从配置表读取要处理的 Sheet 配置信息
 * @param {Spreadsheet} spreadsheet - 表格对象
 * @returns {Map<string, Object>} Sheet 配置信息映射表，key为Sheet名称，value为配置对象
 */
function readSheetConfig(spreadsheet) {
  // 先列出所有 sheet，用于调试
  const allSheets = spreadsheet.getSheets();
  const allSheetNames = allSheets.map(s => s.getName());
  Logger.log(`当前表格中的所有 Sheet: ${allSheetNames.join(', ')}`);
  Logger.log(`正在查找配置表: ${CONFIG.CONFIG_SHEET_NAME}`);
  
  const configSheet = spreadsheet.getSheetByName(CONFIG.CONFIG_SHEET_NAME);
  
  // 如果配置表不存在，直接报错
  if (!configSheet) {
    throw new Error(`配置表 ${CONFIG.CONFIG_SHEET_NAME} 不存在，请先创建配置表`);
  }
  
  Logger.log(`✓ 找到配置表: ${CONFIG.CONFIG_SHEET_NAME}`);
  
  // 读取配置表数据
  const dataRange = configSheet.getDataRange();
  const values = dataRange.getValues();
  
  Logger.log(`配置表数据行数: ${values.length}`);
  
  if (values.length < 2) {
    throw new Error(`配置表 ${CONFIG.CONFIG_SHEET_NAME} 没有数据（只有表头），请至少添加一行数据`);
  }
  
  // 解析表头，找到"Sheet名称"和"启用状态"列
  const headers = values[0];
  Logger.log(`配置表表头: ${headers.join(', ')}`);
  
  const headerMap = {};
  headers.forEach((header, index) => {
    const normalizedHeader = String(header).trim().toLowerCase();
    headerMap[normalizedHeader] = index;
    Logger.log(`  表头[${index}]: "${header}" -> 标准化: "${normalizedHeader}"`);
  });
  
  // 支持多种表头名称（更灵活的匹配）
  // 注意：headerMap 中的键都是小写的，所以查找时也要用小写
  let sheetNameHeader = headerMap['sheet名称'] || 
                        headerMap['sheet name'] || 
                        headerMap['名称'] || 
                        headerMap['name'] || 
                        headerMap['sheet'] || 
                        headerMap['表名'] ||
                        headerMap['工作表名称'] ||
                        headerMap['工作表'] ||
                        headerMap['tab名称'] ||
                        headerMap['tab name'];
  
  // 如果还没找到，尝试更宽松的匹配：遍历所有键，查找包含"sheet"或"名称"的键
  if (sheetNameHeader === undefined) {
    for (const key of Object.keys(headerMap)) {
      if (key.includes('sheet') || key.includes('名称') || key === 'name' || key === '表名' || key.includes('工作表')) {
        sheetNameHeader = headerMap[key];
        Logger.log(`通过宽松匹配找到Sheet名称列: "${key}" (索引: ${sheetNameHeader})`);
        break;
      }
    }
  }
  
  const enabledHeader = headerMap['启用状态'] || 
                        headerMap['enabled'] || 
                        headerMap['启用'] || 
                        headerMap['状态'] || 
                        headerMap['status'] || 
                        headerMap['是否启用'] ||
                        headerMap['enable'] ||
                        headerMap['active'];
  
  // 读取邮箱和日历ID列
  const teacherCalendarIdHeader = headerMap['老师日历授权id'] || 
                                   headerMap['teacher calendar id'] || 
                                   headerMap['老师日历id'] ||
                                   headerMap['teachercalendarid'] ||
                                   headerMap['老师日历授权id'] ||
                                   headerMap['teacher calendar id'];
  
  const studentCalendarIdHeader = headerMap['学生日历授权id'] || 
                                   headerMap['student calendar id'] || 
                                   headerMap['学生日历id'] ||
                                   headerMap['studentcalendarid'] ||
                                   headerMap['学生日历授权id'] ||
                                   headerMap['student calendar id'];
  
  const teacherEmailHeader = headerMap['老师邮箱'] || 
                             headerMap['teacher email'] || 
                             headerMap['老师email'] ||
                             headerMap['teacheremail'] ||
                             headerMap['老师邮件'];
  
  const studentEmailHeader = headerMap['学生邮箱'] || 
                             headerMap['student email'] || 
                             headerMap['学生email'] ||
                             headerMap['studentemail'] ||
                             headerMap['学生邮件'];
  
  const timezoneHeader = headerMap['时区'] || 
                         headerMap['timezone'] || 
                         headerMap['time zone'] ||
                         headerMap['tz'];
  
  const reminderMinutesHeader = headerMap['提醒时间'] || 
                                headerMap['reminder minutes'] || 
                                headerMap['reminder'] ||
                                headerMap['提醒'] ||
                                headerMap['邮件提醒'] ||
                                headerMap['email reminder'] ||
                                headerMap['提前提醒'] ||
                                headerMap['minutes before'];
  
  Logger.log(`Sheet名称列索引: ${sheetNameHeader !== undefined ? sheetNameHeader : '未找到'}`);
  Logger.log(`启用状态列索引: ${enabledHeader !== undefined ? enabledHeader : '未找到'}`);
  Logger.log(`老师日历授权ID列索引: ${teacherCalendarIdHeader !== undefined ? teacherCalendarIdHeader : '未找到'}`);
  Logger.log(`学生日历授权ID列索引: ${studentCalendarIdHeader !== undefined ? studentCalendarIdHeader : '未找到'}`);
  Logger.log(`老师邮箱列索引: ${teacherEmailHeader !== undefined ? teacherEmailHeader : '未找到'}`);
  Logger.log(`学生邮箱列索引: ${studentEmailHeader !== undefined ? studentEmailHeader : '未找到'}`);
  Logger.log(`时区列索引: ${timezoneHeader !== undefined ? timezoneHeader : '未找到'}`);
  Logger.log(`提醒时间列索引: ${reminderMinutesHeader !== undefined ? reminderMinutesHeader : '未找到'}`);
  
  if (sheetNameHeader === undefined) {
    const availableHeaders = Object.keys(headerMap).join(', ');
    throw new Error(`配置表 ${CONFIG.CONFIG_SHEET_NAME} 缺少"Sheet名称"列。\n当前表头: ${headers.join(', ')}\n可用的表头键: ${availableHeaders}\n请确保包含 Sheet 名称的列，支持的列名：Sheet名称、Sheet Name、名称、Name、Sheet、表名等`);
  }
  
  // 读取启用的 Sheet 配置信息
  const sheetConfigMap = new Map();
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const sheetName = row[sheetNameHeader];
    
    Logger.log(`读取第 ${i + 1} 行: Sheet名称="${sheetName}"`);
    
    // 跳过空行
    if (!sheetName || String(sheetName).trim() === '') {
      Logger.log(`  跳过空行`);
      continue;
    }
    
    const sheetNameTrimmed = String(sheetName).trim();
    
    // 检查启用状态（如果存在启用状态列）
    if (enabledHeader !== undefined) {
      const enabled = row[enabledHeader];
      const enabledStr = String(enabled).trim().toLowerCase();
      Logger.log(`  启用状态: "${enabled}" -> 标准化: "${enabledStr}"`);
      // 支持多种表示方式：是/Yes/1/true/启用
      if (enabledStr !== '是' && enabledStr !== 'yes' && enabledStr !== '1' && enabledStr !== 'true' && enabledStr !== '启用' && enabledStr !== 'enabled') {
        Logger.log(`  跳过未启用的 Sheet: ${sheetNameTrimmed}`);
        continue;
      }
    } else {
      Logger.log(`  未找到启用状态列，默认启用`);
    }
    
    // 验证 Sheet 是否存在
    const sheet = spreadsheet.getSheetByName(sheetNameTrimmed);
    if (!sheet) {
      Logger.log(`  警告：配置的 Sheet "${sheetNameTrimmed}" 不存在，已跳过`);
      Logger.log(`  当前所有 Sheet: ${allSheetNames.join(', ')}`);
      continue;
    }
    
    // 读取配置信息
    // 确保提醒时间字段是字符串类型
    let reminderMinutesStr = '';
    if (reminderMinutesHeader !== undefined && row[reminderMinutesHeader] !== undefined && row[reminderMinutesHeader] !== null && row[reminderMinutesHeader] !== '') {
      reminderMinutesStr = String(row[reminderMinutesHeader]).trim();
    }
    
    let reminderMinutes = null;
    
    // 解析提醒时间（支持分钟数，如：30、60、120等）
    if (reminderMinutesStr) {
      const parsed = parseInt(reminderMinutesStr, 10);
      if (!isNaN(parsed) && parsed > 0) {
        reminderMinutes = parsed;
      } else {
        Logger.log(`  警告：提醒时间格式不正确，将忽略: "${reminderMinutesStr}"`);
      }
    }
    
    const config = {
      sheetName: sheetNameTrimmed,
      teacherCalendarId: teacherCalendarIdHeader !== undefined ? (row[teacherCalendarIdHeader] || '').trim() : '',
      studentCalendarId: studentCalendarIdHeader !== undefined ? (row[studentCalendarIdHeader] || '').trim() : '',
      teacherEmail: teacherEmailHeader !== undefined ? (row[teacherEmailHeader] || '').trim() : '',
      studentEmail: studentEmailHeader !== undefined ? (row[studentEmailHeader] || '').trim() : '',
      timezone: timezoneHeader !== undefined ? (row[timezoneHeader] || '').trim() : CONFIG.TIMEZONE,
      reminderMinutes: reminderMinutes
    };
    
    // 如果邮箱为空，尝试使用日历ID作为邮箱
    if (!config.teacherEmail && config.teacherCalendarId) {
      config.teacherEmail = config.teacherCalendarId;
    }
    if (!config.studentEmail && config.studentCalendarId) {
      config.studentEmail = config.studentCalendarId;
    }
    
    // 如果时区为空，使用默认时区
    if (!config.timezone) {
      config.timezone = CONFIG.TIMEZONE;
    }
    
    Logger.log(`  ✓ 添加 Sheet: ${sheetNameTrimmed}`);
    Logger.log(`    老师日历ID: ${config.teacherCalendarId}, 老师邮箱: ${config.teacherEmail}`);
    Logger.log(`    学生日历ID: ${config.studentCalendarId}, 学生邮箱: ${config.studentEmail}`);
    Logger.log(`    时区: ${config.timezone}`);
    Logger.log(`    提醒时间: ${config.reminderMinutes ? config.reminderMinutes + '分钟' : '未配置'}`);
    
    sheetConfigMap.set(sheetNameTrimmed, config);
  }
  
  Logger.log(`从配置表读取到 ${sheetConfigMap.size} 个启用的 Sheet 配置`);
  return sheetConfigMap;
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
            // 添加延迟，避免速率限制
            addOperationDelay();
          } catch (error) {
            Logger.log(`删除旧老师日历事件失败: ${oldRecord.teacherEventId} - ${error.message}`);
            // 如果是速率限制错误，记录详细信息
            if (isRateLimitError(error)) {
              Logger.log(`⚠️ 删除旧老师日历事件遇到速率限制，可能需要稍后重试`);
            }
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
            // 添加延迟，避免速率限制
            addOperationDelay();
          } catch (error) {
            Logger.log(`删除旧学生日历事件失败: ${oldRecord.studentEventId} - ${error.message}`);
            // 如果是速率限制错误，记录详细信息
            if (isRateLimitError(error)) {
              Logger.log(`⚠️ 删除旧学生日历事件遇到速率限制，可能需要稍后重试`);
            }
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
        if (teacherEventId) {
          result.teacherEmail.eventId = String(teacherEventId);
          if (existingInfo.teacherEventId && existingInfo.hasChanges) {
            Logger.log(`老师日历事件更新成功: ${teacherEventId}`);
          } else if (existingInfo.teacherEventId) {
            Logger.log(`老师日历事件保持不变: ${teacherEventId}`);
          } else {
            Logger.log(`老师日历事件创建成功: ${teacherEventId}`);
          }
        } else {
          result.teacherEmail.error = '创建事件成功但未返回事件ID';
          Logger.log(`老师日历事件处理失败: 创建事件成功但未返回事件ID`);
        }
        // 添加延迟，避免速率限制
        addOperationDelay();
      } catch (error) {
        result.teacherEmail.error = error.message;
        Logger.log(`老师日历事件处理失败: ${error.message}`);
        // 如果是速率限制错误，记录详细信息
        if (isRateLimitError(error)) {
          Logger.log(`⚠️ 老师日历事件遇到速率限制，可能需要稍后重试`);
        }
        // 即使创建失败，也尝试保留已有的事件ID（如果有）
        if (existingInfo.teacherEventId) {
          result.teacherEmail.eventId = String(existingInfo.teacherEventId);
          Logger.log(`保留已有老师日历事件ID: ${existingInfo.teacherEventId}`);
        }
      }
    } else {
      // token相同且已有事件ID，跳过更新
      result.teacherEmail.eventId = existingInfo.teacherEventId ? String(existingInfo.teacherEventId) : null;
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
        if (studentEventId) {
          result.studentEmail.eventId = String(studentEventId);
          if (existingInfo.studentEventId && existingInfo.hasChanges) {
            Logger.log(`学生日历事件更新成功: ${studentEventId}`);
          } else if (existingInfo.studentEventId) {
            Logger.log(`学生日历事件保持不变: ${studentEventId}`);
          } else {
            Logger.log(`学生日历事件创建成功: ${studentEventId}`);
          }
        } else {
          result.studentEmail.error = '创建事件成功但未返回事件ID';
          Logger.log(`学生日历事件处理失败: 创建事件成功但未返回事件ID`);
        }
        // 添加延迟，避免速率限制
        addOperationDelay();
      } catch (error) {
        result.studentEmail.error = error.message;
        Logger.log(`学生日历事件处理失败: ${error.message}`);
        // 如果是速率限制错误，记录详细信息
        if (isRateLimitError(error)) {
          Logger.log(`⚠️ 学生日历事件遇到速率限制，可能需要稍后重试`);
        }
        // 即使创建失败，也尝试保留已有的事件ID（如果有）
        if (existingInfo.studentEventId) {
          result.studentEmail.eventId = String(existingInfo.studentEventId);
          Logger.log(`保留已有学生日历事件ID: ${existingInfo.studentEventId}`);
        }
      }
    } else {
      // token相同且已有事件ID，跳过更新
      result.studentEmail.eventId = existingInfo.studentEventId ? String(existingInfo.studentEventId) : null;
      Logger.log(`学生日历事件跳过（token相同且已有事件）: ${existingInfo.studentEventId}`);
    }
    
    // 5. 判断整体状态
    // 如果邮件跳过（因为token没变化），不应该影响成功判断
    // 只要日历事件创建成功，就算成功
    const teacherEventId = result.teacherEmail.eventId ? String(result.teacherEmail.eventId).trim() : '';
    const studentEventId = result.studentEmail.eventId ? String(result.studentEmail.eventId).trim() : '';
    const teacherSuccess = teacherEventId !== '' && !result.teacherEmail.error;
    const studentSuccess = studentEventId !== '' && !result.studentEmail.error;
    
    Logger.log(`[${course.lessonNumber}] 状态判断: 老师事件ID=${teacherEventId || '无'}, 学生事件ID=${studentEventId || '无'}, 老师成功=${teacherSuccess}, 学生成功=${studentSuccess}`);
    
    if (teacherSuccess && studentSuccess) {
      result.status = '已完成';
    } else if (teacherSuccess || studentSuccess) {
      result.status = '部分失败';
    } else {
      result.status = '失败';
    }
    
    Logger.log(`[${course.lessonNumber}] 最终状态: ${result.status}`);
    
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
 * @param {Sheet} sheet - 课程表对象
 * @param {Object} config - Sheet 配置信息（包含邮箱和日历ID）
 * @returns {Array<Object>} 课程数据数组
 */
function readCourseData(sheet, config) {
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
        studentName: row[headerMap['学生']] || '',
        startTime: row[headerMap['开始时间']] || '',
        endTime: row[headerMap['结束时间']] || '',
        // 从配置中获取邮箱和日历ID
        teacherEmail: config.teacherEmail || '',
        studentEmail: config.studentEmail || '',
        teacherCalendarId: config.teacherCalendarId || config.teacherEmail || '',
        studentCalendarId: config.studentCalendarId || config.studentEmail || '',
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
  
  // 读取表头，建立表头名称到列索引的映射
  const headers = values[0];
  const headerMap = {};
  headers.forEach((header, index) => {
    const headerKey = String(header).trim().toLowerCase();
    headerMap[headerKey] = index;
  });
  
  // 定义表头名称的多种变体（支持中英文）
  const getColumnIndex = (headerNames) => {
    for (const name of headerNames) {
      const key = name.toLowerCase();
      if (headerMap[key] !== undefined) {
        return headerMap[key];
      }
    }
    return undefined;
  };
  
  // 获取各列的索引（使用表头名称而不是固定索引）
  const recordIdCol = getColumnIndex(['记录id', 'record id', '记录id', 'recordid', 'id']);
  const lessonNumberCol = getColumnIndex(['课次', 'lesson', 'lesson number', '课程次数']);
  const dateCol = getColumnIndex(['日期', 'date', '课程日期']);
  const tokenCol = getColumnIndex(['token', '令牌', '哈希']);
  const teacherEmailStatusCol = getColumnIndex(['老师邮件状态', 'teacher email status', '老师邮件']);
  const teacherEmailTimeCol = getColumnIndex(['老师邮件发送时间', 'teacher email time', '老师邮件时间']);
  const teacherCalendarIdCol = getColumnIndex(['老师日历id', 'teacher calendar id', '老师日历']);
  const teacherEventIdCol = getColumnIndex(['老师日历事件id', 'teacher event id', '老师事件id']);
  const teacherEventTimeCol = getColumnIndex(['老师日历创建时间', 'teacher event time', '老师事件时间']);
  const studentEmailStatusCol = getColumnIndex(['学生邮件状态', 'student email status', '学生邮件']);
  const studentEmailTimeCol = getColumnIndex(['学生邮件发送时间', 'student email time', '学生邮件时间']);
  const studentCalendarIdCol = getColumnIndex(['学生日历id', 'student calendar id', '学生日历']);
  const studentEventIdCol = getColumnIndex(['学生日历事件id', 'student event id', '学生事件id']);
  const studentEventTimeCol = getColumnIndex(['学生日历创建时间', 'student event time', '学生事件时间']);
  const statusCol = getColumnIndex(['处理状态', 'status', '状态']);
  const lastUpdateTimeCol = getColumnIndex(['最后更新时间', 'last update time', '更新时间']);
  
  // 从第2行开始读取（第1行为表头）
  // 状态表的第i行对应正式表的第i行（都有表头）
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    
    // 使用表头映射获取值
    const getValue = (colIndex) => {
      if (colIndex === undefined) return '';
      return row[colIndex] || '';
    };
    
    // 如果课次和日期都为空，跳过（空行）
    const lessonNumber = getValue(lessonNumberCol);
    const date = getValue(dateCol);
    if (!lessonNumber && !date) {
      continue;
    }
    
    const recordId = getValue(recordIdCol);
    const key = `${lessonNumber}_${date}`; // 课次_日期（向后兼容）
    
    const record = {
      recordId: recordId, // 记录ID
      lessonNumber: lessonNumber,
      date: date,
      token: getValue(tokenCol), // Token（关键信息哈希）
      teacherCalendarId: (getValue(teacherCalendarIdCol) && !(getValue(teacherCalendarIdCol) instanceof Date) && String(getValue(teacherCalendarIdCol)).trim()) || '', // 老师日历ID（用于删除事件）
      teacherEventId: (getValue(teacherEventIdCol) && !(getValue(teacherEventIdCol) instanceof Date) && String(getValue(teacherEventIdCol)).trim()) || '', // 老师日历事件ID
      studentCalendarId: (getValue(studentCalendarIdCol) && !(getValue(studentCalendarIdCol) instanceof Date) && String(getValue(studentCalendarIdCol)).trim()) || '', // 学生日历ID（用于删除事件）
      studentEventId: (getValue(studentEventIdCol) && !(getValue(studentEventIdCol) instanceof Date) && String(getValue(studentEventIdCol)).trim()) || '', // 学生日历事件ID
      status: getValue(statusCol), // 处理状态
      rowIndex: i + 1 // 状态表的行号（从1开始，包含表头）
    };
    
    // 验证事件ID格式：如果事件ID是"已发送"或其他状态文本，说明是错误的数据，应该清空
    const invalidStatusTexts = ['已发送', '未发送', '失败', '部分失败', '已完成', '处理中'];
    if (record.teacherEventId && invalidStatusTexts.includes(record.teacherEventId)) {
      Logger.log(`警告：老师事件ID包含状态文本，将被清空: "${record.teacherEventId}"`);
      record.teacherEventId = '';
    }
    if (record.studentEventId && invalidStatusTexts.includes(record.studentEventId)) {
      Logger.log(`警告：学生事件ID包含状态文本，将被清空: "${record.studentEventId}"`);
      record.studentEventId = '';
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
      // 添加延迟，避免速率限制
      addOperationDelay();
    } catch (error) {
      Logger.log(`删除老师日历事件失败: ${deletedRecord.teacherEventId} - ${error.message}`);
      // 如果是速率限制错误，记录详细信息
      if (isRateLimitError(error)) {
        Logger.log(`⚠️ 删除老师日历事件遇到速率限制，可能需要稍后重试`);
      }
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
      // 添加延迟，避免速率限制
      addOperationDelay();
    } catch (error) {
      Logger.log(`删除学生日历事件失败: ${deletedRecord.studentEventId} - ${error.message}`);
      // 如果是速率限制错误，记录详细信息
      if (isRateLimitError(error)) {
        Logger.log(`⚠️ 删除学生日历事件遇到速率限制，可能需要稍后重试`);
      }
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
        deleteEventWithRetry(event);
        Logger.log(`删除日历事件成功: ${eventId} (日历: ${calendar.getName()})`);
        return; // 找到并删除后退出
      }
    } catch (error) {
      // 如果是速率限制错误，记录并继续
      if (isRateLimitError(error)) {
        Logger.log(`删除日历事件时遇到速率限制: ${eventId} - ${error.message}`);
      }
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
  // 使用默认时区格式化日期（取消邮件时可能没有 course 对象）
  const timezone = CONFIG.TIMEZONE || Session.getScriptTimeZone();
  const dateStr = Utilities.formatDate(eventDate, timezone, 'yyyy-MM-dd');
  
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
 * @param {Sheet} statusSheet - 状态表
 * @param {string} lessonNumber - 课次
 * @param {Date|string} currentDate - 当前日期
 * @param {string} timezone - 时区（可选，默认使用脚本时区）
 */
function findOldRecordsByLessonNumber(statusSheet, lessonNumber, currentDate, timezone) {
  const oldRecords = [];
  
  if (!statusSheet || statusSheet.getLastRow() < 2) {
    return oldRecords;
  }
  
  // 获取时区（优先使用传入的时区，否则使用默认时区）
  const tz = timezone || CONFIG.TIMEZONE || Session.getScriptTimeZone();
  
  const dataRange = statusSheet.getDataRange();
  const values = dataRange.getValues();
  
  // 读取表头，建立表头名称到列索引的映射
  const headers = values[0];
  const headerMap = {};
  headers.forEach((header, index) => {
    const headerKey = String(header).trim().toLowerCase();
    headerMap[headerKey] = index;
  });
  
  // 定义表头名称的多种变体（支持中英文）
  const getColumnIndex = (headerNames) => {
    for (const name of headerNames) {
      const key = name.toLowerCase();
      if (headerMap[key] !== undefined) {
        return headerMap[key];
      }
    }
    return undefined;
  };
  
  // 获取各列的索引（使用表头名称而不是固定索引）
  const lessonNumberCol = getColumnIndex(['课次', 'lesson', 'lesson number', '课程次数']);
  const dateCol = getColumnIndex(['日期', 'date', '课程日期']);
  const teacherCalendarIdCol = getColumnIndex(['老师日历id', 'teacher calendar id', '老师日历']);
  const teacherEventIdCol = getColumnIndex(['老师日历事件id', 'teacher event id', '老师事件id']);
  const studentCalendarIdCol = getColumnIndex(['学生日历id', 'student calendar id', '学生日历']);
  const studentEventIdCol = getColumnIndex(['学生日历事件id', 'student event id', '学生事件id']);
  
  // 标准化当前日期用于比较
  const currentDateStr = currentDate instanceof Date ?
    Utilities.formatDate(currentDate, tz, 'yyyy-MM-dd') :
    String(currentDate);
  
  // 使用表头映射获取值
  const getValue = (row, colIndex) => {
    if (colIndex === undefined) return '';
    return row[colIndex] || '';
  };
  
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowLessonNumber = getValue(row, lessonNumberCol);
    const rowDate = getValue(row, dateCol);
    
    // 如果课次相同但日期不同
    if (rowLessonNumber === lessonNumber && rowDate) {
      const rowDateStr = rowDate instanceof Date ?
        Utilities.formatDate(rowDate, tz, 'yyyy-MM-dd') :
        String(rowDate);
      
      if (rowDateStr !== currentDateStr) {
        oldRecords.push({
          lessonNumber: rowLessonNumber,
          date: rowDate,
          teacherCalendarId: getValue(row, teacherCalendarIdCol),
          teacherEventId: getValue(row, teacherEventIdCol),
          studentCalendarId: getValue(row, studentCalendarIdCol),
          studentEventId: getValue(row, studentEventIdCol),
          rowIndex: i + 1
        });
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
 * 验证日历事件是否存在
 * @param {string} calendarId - 日历ID
 * @param {string} eventId - 事件ID
 * @returns {boolean} 事件是否存在
 */
function verifyCalendarEventExists(calendarId, eventId) {
  if (!calendarId || !eventId) {
    return false;
  }
  
  try {
    // 使用更健壮的获取日历方法
    const calendar = getCalendarByIdOrEmail(calendarId, null);
    if (!calendar) {
      Logger.log(`验证事件时找不到日历: ${calendarId}`);
      return false;
    }
    
    // 尝试获取事件
    const event = calendar.getEventById(eventId);
    if (event) {
      Logger.log(`✓ 验证事件存在: ${eventId} (日历: ${calendarId})`);
      return true;
    } else {
      Logger.log(`✗ 验证事件不存在: ${eventId} (日历: ${calendarId})`);
      return false;
    }
  } catch (error) {
    // 如果获取事件时抛出异常，通常表示事件不存在
    Logger.log(`验证事件时出错: ${eventId} (日历: ${calendarId}) - ${error.message}`);
    return false;
  }
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
      deleteEventWithRetry(event);
      Logger.log(`删除日历事件成功: ${eventId} (日历: ${calendarId})`);
    } else {
      Logger.log(`找不到日历事件: ${eventId} (日历: ${calendarId})`);
    }
  } catch (error) {
    Logger.log(`删除日历事件失败: ${eventId} (日历: ${calendarId}) - ${error.message}`);
    // 如果是速率限制错误，抛出异常以便上层处理
    if (isRateLimitError(error)) {
      throw error;
    }
  }
}

/**
 * 计算课程关键信息的token（用于检测变化）
 * 包括：日期、开始时间、结束时间、课程内容、老师、老师邮箱、学生、学生邮箱
 * @param {Object} course - 课程对象，包含 timezone 属性
 */
function calculateCourseToken(course) {
  // 获取时区（优先使用课程配置的时区，否则使用默认时区）
  const timezone = course.timezone || CONFIG.TIMEZONE || Session.getScriptTimeZone();
  
  // 标准化日期和时间格式
  const dateStr = course.date instanceof Date ? 
    Utilities.formatDate(course.date, timezone, 'yyyy-MM-dd') : 
    String(course.date);
  
  const startTimeStr = course.startTime instanceof Date ?
    Utilities.formatDate(course.startTime, timezone, 'HH:mm') :
    String(course.startTime);
  
  const endTimeStr = course.endTime instanceof Date ?
    Utilities.formatDate(course.endTime, timezone, 'HH:mm') :
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

// ==================== 速率限制处理模块 ====================

/**
 * 检查是否是速率限制错误
 * @param {Error} error - 错误对象
 * @returns {boolean} 是否是速率限制错误
 */
function isRateLimitError(error) {
  if (!error || !error.message) {
    return false;
  }
  
  const errorMessage = error.message.toLowerCase();
  return CONFIG.RATE_LIMIT.RATE_LIMIT_KEYWORDS.some(keyword => 
    errorMessage.includes(keyword.toLowerCase())
  );
}

/**
 * 带重试的创建日历事件
 * @param {Calendar} calendar - 日历对象
 * @param {string} title - 事件标题
 * @param {Date} startTime - 开始时间
 * @param {Date} endTime - 结束时间
 * @param {Object} options - 选项（description, guests, sendInvites）
 * @returns {CalendarEvent} 创建的事件对象
 */
function createEventWithRetry(calendar, title, startTime, endTime, options) {
  let lastError;
  const maxRetries = CONFIG.RATE_LIMIT.MAX_RETRIES;
  const retryDelay = CONFIG.RATE_LIMIT.RETRY_DELAY;
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      // 添加延迟（除了第一次尝试）
      if (attempt > 1) {
        Logger.log(`重试创建日历事件（第${attempt}次尝试）...`);
        Utilities.sleep(retryDelay * (attempt - 1)); // 递增延迟
      }
      
      const event = calendar.createEvent(title, startTime, endTime, options);
      return event;
    } catch (error) {
      lastError = error;
      
      if (isRateLimitError(error)) {
        Logger.log(`遇到速率限制错误（第${attempt}次尝试）: ${error.message}`);
        if (attempt < maxRetries) {
          Logger.log(`等待 ${retryDelay * attempt} 毫秒后重试...`);
          continue;
        } else {
          Logger.log(`已达到最大重试次数（${maxRetries}），放弃创建事件`);
          throw new Error(`创建日历事件失败（速率限制）: ${error.message}`);
        }
      } else {
        // 非速率限制错误，直接抛出
        throw error;
      }
    }
  }
  
  throw lastError || new Error('创建日历事件失败');
}

/**
 * 带重试的删除日历事件
 * @param {CalendarEvent} event - 事件对象
 */
function deleteEventWithRetry(event) {
  let lastError;
  const maxRetries = CONFIG.RATE_LIMIT.MAX_RETRIES;
  const retryDelay = CONFIG.RATE_LIMIT.RETRY_DELAY;
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      // 添加延迟（除了第一次尝试）
      if (attempt > 1) {
        Logger.log(`重试删除日历事件（第${attempt}次尝试）...`);
        Utilities.sleep(retryDelay * (attempt - 1)); // 递增延迟
      }
      
      event.deleteEvent();
      return;
    } catch (error) {
      lastError = error;
      
      if (isRateLimitError(error)) {
        Logger.log(`遇到速率限制错误（第${attempt}次尝试）: ${error.message}`);
        if (attempt < maxRetries) {
          Logger.log(`等待 ${retryDelay * attempt} 毫秒后重试...`);
          continue;
        } else {
          Logger.log(`已达到最大重试次数（${maxRetries}），放弃删除事件`);
          throw new Error(`删除日历事件失败（速率限制）: ${error.message}`);
        }
      } else {
        // 非速率限制错误，直接抛出
        throw error;
      }
    }
  }
  
  throw lastError || new Error('删除日历事件失败');
}

/**
 * 带重试的更新日历事件
 * @param {CalendarEvent} event - 事件对象
 * @param {string} title - 事件标题
 * @param {string} description - 事件描述
 * @param {Date} startTime - 开始时间
 * @param {Date} endTime - 结束时间
 * @param {string} guests - 参与者列表（逗号分隔）
 */
function updateEventWithRetry(event, title, description, startTime, endTime, guests) {
  let lastError;
  const maxRetries = CONFIG.RATE_LIMIT.MAX_RETRIES;
  const retryDelay = CONFIG.RATE_LIMIT.RETRY_DELAY;
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      // 添加延迟（除了第一次尝试）
      if (attempt > 1) {
        Logger.log(`重试更新日历事件（第${attempt}次尝试）...`);
        Utilities.sleep(retryDelay * (attempt - 1)); // 递增延迟
      }
      
      // 更新事件信息
      event.setTitle(title);
      event.setDescription(description);
      event.setTime(startTime, endTime);
      
      // 更新参与者（使用正确的方法）
      // 先获取现有参与者列表
      const existingGuests = event.getGuestList();
      const existingEmails = existingGuests.map(guest => guest.getEmail());
      const newEmails = guests.split(',').map(email => email.trim());
      
      // 添加新参与者
      for (const email of newEmails) {
        if (email && !existingEmails.includes(email)) {
          event.addGuest(email);
        }
      }
      
      // 移除不在新列表中的参与者（可选，根据需求决定）
      // 这里不删除，只添加新的参与者
      
      return;
    } catch (error) {
      lastError = error;
      
      if (isRateLimitError(error)) {
        Logger.log(`遇到速率限制错误（第${attempt}次尝试）: ${error.message}`);
        if (attempt < maxRetries) {
          Logger.log(`等待 ${retryDelay * attempt} 毫秒后重试...`);
          continue;
        } else {
          Logger.log(`已达到最大重试次数（${maxRetries}），放弃更新事件`);
          throw new Error(`更新日历事件失败（速率限制）: ${error.message}`);
        }
      } else {
        // 非速率限制错误，直接抛出
        throw error;
      }
    }
  }
  
  throw lastError || new Error('更新日历事件失败');
}

/**
 * 添加操作延迟（用于避免速率限制）
 */
function addOperationDelay() {
  Utilities.sleep(CONFIG.RATE_LIMIT.DELAY_BETWEEN_OPERATIONS);
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
  
  // 解析日期和时间（使用时区）
  const timezone = course.timezone || CONFIG.TIMEZONE || Session.getScriptTimeZone();
  const startDateTime = parseDateTime(course.date, course.startTime, timezone);
  const endDateTime = parseDateTime(course.date, course.endTime, timezone);
  
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
      
      // 更新事件信息（带速率限制处理）
      updateEventWithRetry(event, eventSummary, eventDescription, eventStart, eventEnd, eventGuests);
      
      // 更新提醒（如果配置了提醒时间）
      if (course.reminderMinutes && course.reminderMinutes > 0) {
        try {
          // 清除现有提醒
          event.removeAllReminders();
          // 添加新的提醒
          event.addEmailReminder(course.reminderMinutes);
          Logger.log(`更新邮件提醒: 提前 ${course.reminderMinutes} 分钟`);
        } catch (error) {
          Logger.log(`更新提醒失败: ${error.message}`);
          // 提醒失败不影响事件更新，继续执行
        }
      } else {
        // 如果没有配置提醒时间，清除现有提醒
        try {
          event.removeAllReminders();
          Logger.log(`清除提醒（未配置提醒时间）`);
        } catch (error) {
          Logger.log(`清除提醒失败: ${error.message}`);
        }
      }
      
      Logger.log(`更新日历事件: ${existingEventId}`);
      return existingEventId;
    } catch (error) {
      // 如果事件不存在或无法访问，则创建新事件
      Logger.log(`无法更新事件 ${existingEventId}，将创建新事件: ${error.message}`);
      // 继续执行创建逻辑
    }
  }
  
  // 创建新事件（带速率限制处理）
  event = createEventWithRetry(
    calendar,
    eventSummary,
    eventStart,
    eventEnd,
    {
      description: eventDescription,
      guests: eventGuests,
      sendInvites: true
    }
  );
  
  // 添加提醒（如果配置了提醒时间）
  if (course.reminderMinutes && course.reminderMinutes > 0) {
    try {
      event.addEmailReminder(course.reminderMinutes);
      Logger.log(`添加邮件提醒: 提前 ${course.reminderMinutes} 分钟`);
    } catch (error) {
      Logger.log(`添加提醒失败: ${error.message}`);
      // 提醒失败不影响事件创建，继续执行
    }
  }
  
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
 * @param {Spreadsheet} spreadsheet - 表格对象
 * @param {string} statusSheetName - 状态表名称（可选，如果不提供则使用默认名称）
 * @returns {Sheet} 状态表对象
 */
function ensureStatusSheet(spreadsheet, statusSheetName) {
  // 如果没有提供状态表名称，使用默认名称（向后兼容）
  const targetStatusSheetName = statusSheetName || CONFIG.STATUS_SHEET_PREFIX + CONFIG.MAIN_SHEET_NAME;
  
  let statusSheet = spreadsheet.getSheetByName(targetStatusSheetName);
  
  if (!statusSheet) {
    // 创建隐藏表
    statusSheet = spreadsheet.insertSheet(targetStatusSheetName);
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
    
    Logger.log(`创建状态表: ${targetStatusSheetName}`);
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
  // 使用课程配置的时区，如果没有则使用默认时区
  const timezone = course.timezone || CONFIG.TIMEZONE || Session.getScriptTimeZone();
  const nowStr = Utilities.formatDate(now, timezone, 'yyyy-MM-dd HH:mm:ss');
  
  // 使用course.rowIndex来确定状态表的行号
  // 状态表的第i行对应正式表的第i+1行（正式表有表头，状态表也有表头）
  const rowIndex = course.rowIndex; // course.rowIndex是正式表的行号（从1开始，包含表头）
  
  // 读取表头，建立表头名称到列索引的映射
  const headerRow = statusSheet.getRange(1, 1, 1, statusSheet.getLastColumn()).getValues()[0];
  const headerMap = {};
  headerRow.forEach((header, index) => {
    const headerKey = String(header).trim().toLowerCase();
    headerMap[headerKey] = index;
  });
  
  // 定义表头名称的多种变体（支持中英文）
  const getColumnIndex = (headerNames) => {
    for (const name of headerNames) {
      const key = name.toLowerCase();
      if (headerMap[key] !== undefined) {
        return headerMap[key];
      }
    }
    return undefined;
  };
  
  // 获取各列的索引（使用表头名称而不是固定索引）
  const recordIdCol = getColumnIndex(['记录id', 'record id', '记录id', 'recordid', 'id']);
  const lessonNumberCol = getColumnIndex(['课次', 'lesson', 'lesson number', '课程次数']);
  const dateCol = getColumnIndex(['日期', 'date', '课程日期']);
  const tokenCol = getColumnIndex(['token', '令牌', '哈希']);
  const teacherEmailStatusCol = getColumnIndex(['老师邮件状态', 'teacher email status', '老师邮件']);
  const teacherEmailTimeCol = getColumnIndex(['老师邮件发送时间', 'teacher email time', '老师邮件时间']);
  const teacherCalendarIdCol = getColumnIndex(['老师日历id', 'teacher calendar id', '老师日历']);
  const teacherEventIdCol = getColumnIndex(['老师日历事件id', 'teacher event id', '老师事件id']);
  const teacherEventTimeCol = getColumnIndex(['老师日历创建时间', 'teacher event time', '老师事件时间']);
  const studentEmailStatusCol = getColumnIndex(['学生邮件状态', 'student email status', '学生邮件']);
  const studentEmailTimeCol = getColumnIndex(['学生邮件发送时间', 'student email time', '学生邮件时间']);
  const studentCalendarIdCol = getColumnIndex(['学生日历id', 'student calendar id', '学生日历']);
  const studentEventIdCol = getColumnIndex(['学生日历事件id', 'student event id', '学生事件id']);
  const studentEventTimeCol = getColumnIndex(['学生日历创建时间', 'student event time', '学生事件时间']);
  const statusCol = getColumnIndex(['处理状态', 'status', '状态']);
  const lastUpdateTimeCol = getColumnIndex(['最后更新时间', 'last update time', '更新时间']);
  
  // 读取当前行的现有记录（如果有）
  let existingRecord = null;
  if (rowIndex <= statusSheet.getLastRow()) {
    const rowValues = statusSheet.getRange(rowIndex, 1, 1, statusSheet.getLastColumn()).getValues()[0];
    // 使用表头映射获取值
    const getValue = (colIndex) => {
      if (colIndex === undefined) return '';
      return rowValues[colIndex] || '';
    };
    // 如果课次或日期不为空，说明有记录
    if (getValue(lessonNumberCol) || getValue(dateCol)) {
      existingRecord = { rowValues, getValue };
    }
  }
  
  // 获取或生成记录ID
  const getExistingValue = (colIndex) => {
    if (!existingRecord || colIndex === undefined) return '';
    return existingRecord.getValue(colIndex);
  };
  const recordId = course.recordId || (existingRecord ? (getExistingValue(recordIdCol) || generateRecordId()) : generateRecordId());
  
  // 保留已有的事件ID和日历ID（如果更新失败）
  // 确保从 existingRecord 中读取的值是字符串
  let existingTeacherCalendarId = getExistingValue(teacherCalendarIdCol);
  existingTeacherCalendarId = existingTeacherCalendarId && !(existingTeacherCalendarId instanceof Date) ? String(existingTeacherCalendarId).trim() : '';
  let existingTeacherEventId = getExistingValue(teacherEventIdCol);
  existingTeacherEventId = existingTeacherEventId && !(existingTeacherEventId instanceof Date) ? String(existingTeacherEventId).trim() : '';
  let existingStudentCalendarId = getExistingValue(studentCalendarIdCol);
  existingStudentCalendarId = existingStudentCalendarId && !(existingStudentCalendarId instanceof Date) ? String(existingStudentCalendarId).trim() : '';
  let existingStudentEventId = getExistingValue(studentEventIdCol);
  existingStudentEventId = existingStudentEventId && !(existingStudentEventId instanceof Date) ? String(existingStudentEventId).trim() : '';
  
  // 验证事件ID格式：如果事件ID是"已发送"或其他状态文本，说明是错误的数据，应该清空
  const invalidStatusTexts = ['已发送', '未发送', '失败', '部分失败', '已完成', '处理中'];
  if (existingTeacherEventId && invalidStatusTexts.includes(existingTeacherEventId)) {
    Logger.log(`警告：老师事件ID包含状态文本，将被清空: "${existingTeacherEventId}"`);
    existingTeacherEventId = '';
  }
  if (existingStudentEventId && invalidStatusTexts.includes(existingStudentEventId)) {
    Logger.log(`警告：学生事件ID包含状态文本，将被清空: "${existingStudentEventId}"`);
    existingStudentEventId = '';
  }
  
  const teacherCalendarId = course.teacherCalendarId || existingTeacherCalendarId || '';
  // 确保事件ID是字符串格式，且不是日期对象或状态文本
  let teacherEventId = '';
  if (result.teacherEmail.eventId) {
    const eventId = result.teacherEmail.eventId;
    // 检查是否是日期对象
    if (eventId instanceof Date) {
      Logger.log(`警告：老师事件ID是日期对象，将被忽略: ${eventId}`);
      teacherEventId = existingTeacherEventId || '';
    } else {
      const eventIdStr = String(eventId).trim();
      // 验证事件ID格式：如果事件ID是"已发送"或其他状态文本，说明是错误的数据，应该清空
      const invalidStatusTexts = ['已发送', '未发送', '失败', '部分失败', '已完成', '处理中'];
      if (invalidStatusTexts.includes(eventIdStr)) {
        Logger.log(`警告：老师事件ID包含状态文本，将被忽略: "${eventIdStr}"`);
        teacherEventId = existingTeacherEventId || '';
      } else {
        teacherEventId = eventIdStr;
      }
    }
  } else {
    teacherEventId = existingTeacherEventId || '';
  }
  
  const studentCalendarId = course.studentCalendarId || existingStudentCalendarId || '';
  // 确保事件ID是字符串格式，且不是日期对象或状态文本
  let studentEventId = '';
  if (result.studentEmail.eventId) {
    const eventId = result.studentEmail.eventId;
    // 检查是否是日期对象
    if (eventId instanceof Date) {
      Logger.log(`警告：学生事件ID是日期对象，将被忽略: ${eventId}`);
      studentEventId = existingStudentEventId || '';
    } else {
      const eventIdStr = String(eventId).trim();
      // 验证事件ID格式：如果事件ID是"已发送"或其他状态文本，说明是错误的数据，应该清空
      const invalidStatusTexts = ['已发送', '未发送', '失败', '部分失败', '已完成', '处理中'];
      if (invalidStatusTexts.includes(eventIdStr)) {
        Logger.log(`警告：学生事件ID包含状态文本，将被忽略: "${eventIdStr}"`);
        studentEventId = existingStudentEventId || '';
      } else {
        studentEventId = eventIdStr;
      }
    }
  } else {
    studentEventId = existingStudentEventId || '';
  }
  
  // 如果事件ID存在，更新创建时间；如果是新创建的，使用当前时间；如果是已有的，保留原时间
  let teacherEventTime = '';
  let studentEventTime = '';
  
  if (result.teacherEmail.eventId && !(result.teacherEmail.eventId instanceof Date)) {
    // 新创建或更新的事件
    teacherEventTime = nowStr;
  } else if (existingRecord && existingTeacherEventId) {
    // 保留原有的创建时间
    const existingTime = getExistingValue(teacherEventTimeCol);
    if (existingTime instanceof Date) {
      // 如果是日期对象，格式化为字符串
      const timezone = course.timezone || CONFIG.TIMEZONE || Session.getScriptTimeZone();
      teacherEventTime = Utilities.formatDate(existingTime, timezone, 'yyyy-MM-dd HH:mm:ss');
    } else if (existingTime) {
      teacherEventTime = String(existingTime).trim();
    }
  }
  
  if (result.studentEmail.eventId && !(result.studentEmail.eventId instanceof Date)) {
    // 新创建或更新的事件
    studentEventTime = nowStr;
  } else if (existingRecord && existingStudentEventId) {
    // 保留原有的创建时间
    const existingTime = getExistingValue(studentEventTimeCol);
    if (existingTime instanceof Date) {
      // 如果是日期对象，格式化为字符串
      const timezone = course.timezone || CONFIG.TIMEZONE || Session.getScriptTimeZone();
      studentEventTime = Utilities.formatDate(existingTime, timezone, 'yyyy-MM-dd HH:mm:ss');
    } else if (existingTime) {
      studentEventTime = String(existingTime).trim();
    }
  }
  
  // 获取或计算token
  const token = course.token || calculateCourseToken(course);
  
  // 格式化日期（确保是字符串格式）
  const dateStr = course.date instanceof Date ? 
    Utilities.formatDate(course.date, course.timezone || CONFIG.TIMEZONE || Session.getScriptTimeZone(), 'yyyy-MM-dd') : 
    String(course.date);
  
  // 使用表头映射来写入数据，而不是固定的列索引
  // 获取所有列索引，确保列存在
  const allColumns = [
    recordIdCol, lessonNumberCol, dateCol, tokenCol,
    teacherEmailStatusCol, teacherEmailTimeCol, teacherCalendarIdCol, teacherEventIdCol, teacherEventTimeCol,
    studentEmailStatusCol, studentEmailTimeCol, studentCalendarIdCol, studentEventIdCol, studentEventTimeCol,
    statusCol, lastUpdateTimeCol
  ];
  
  // 找到最大列索引，确定需要写入的列数
  const maxColIndex = Math.max(...allColumns.filter(col => col !== undefined));
  const totalCols = maxColIndex + 1;
  
  // 创建一行数据数组，初始化为空字符串
  const rowData = new Array(totalCols).fill('');
  
  // 根据表头映射写入数据
  if (recordIdCol !== undefined) rowData[recordIdCol] = recordId;
  if (lessonNumberCol !== undefined) rowData[lessonNumberCol] = course.lessonNumber;
  if (dateCol !== undefined) rowData[dateCol] = dateStr;
  if (tokenCol !== undefined) rowData[tokenCol] = token;
  if (teacherEmailStatusCol !== undefined) {
    rowData[teacherEmailStatusCol] = result.teacherEmail.sent ? '已发送' : (result.teacherEmail.error || (existingRecord ? getExistingValue(teacherEmailStatusCol) : '未发送'));
  }
  if (teacherEmailTimeCol !== undefined) {
    rowData[teacherEmailTimeCol] = result.teacherEmail.sent ? nowStr : (existingRecord ? getExistingValue(teacherEmailTimeCol) : '');
  }
  if (teacherCalendarIdCol !== undefined) rowData[teacherCalendarIdCol] = String(teacherCalendarId || '');
  if (teacherEventIdCol !== undefined) rowData[teacherEventIdCol] = String(teacherEventId || '');
  if (teacherEventTimeCol !== undefined) rowData[teacherEventTimeCol] = String(teacherEventTime || '');
  if (studentEmailStatusCol !== undefined) {
    rowData[studentEmailStatusCol] = result.studentEmail.sent ? '已发送' : (result.studentEmail.error || (existingRecord ? getExistingValue(studentEmailStatusCol) : '未发送'));
  }
  if (studentEmailTimeCol !== undefined) {
    rowData[studentEmailTimeCol] = result.studentEmail.sent ? nowStr : (existingRecord ? getExistingValue(studentEmailTimeCol) : '');
  }
  if (studentCalendarIdCol !== undefined) rowData[studentCalendarIdCol] = String(studentCalendarId || '');
  if (studentEventIdCol !== undefined) rowData[studentEventIdCol] = String(studentEventId || '');
  if (studentEventTimeCol !== undefined) rowData[studentEventTimeCol] = String(studentEventTime || '');
  if (statusCol !== undefined) rowData[statusCol] = result.status;
  if (lastUpdateTimeCol !== undefined) rowData[lastUpdateTimeCol] = nowStr;
  
  // 直接更新对应行（状态表和正式表一一对应）
  statusSheet.getRange(rowIndex, 1, 1, totalCols).setValues([rowData]);
}

// ==================== 工具函数 ====================

/**
 * 解析日期时间
 */
/**
 * 解析日期时间
 * @param {Date|string} dateInput - 日期输入
 * @param {Date|string|number} timeInput - 时间输入
 * @param {string} timezone - 时区（可选，默认使用脚本时区）
 * @returns {Date} 解析后的日期时间对象
 */
function parseDateTime(dateInput, timeInput, timezone) {
  try {
    // 获取时区（优先使用传入的时区，否则使用默认时区）
    const tz = timezone || CONFIG.TIMEZONE || Session.getScriptTimeZone();
    
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
        // 使用指定时区创建日期
        const dateStr = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        date = new Date(dateStr + 'T00:00:00');
      } else if (dateInput.includes('-')) {
        date = new Date(dateInput + 'T00:00:00');
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
    
    // 设置时间（使用指定时区）
    // 先构建日期时间字符串，然后使用时区解析
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const hourStr = String(hours).padStart(2, '0');
    const minuteStr = String(minutes).padStart(2, '0');
    const secondStr = String(seconds).padStart(2, '0');
    
    // 构建日期时间字符串（指定时区的本地时间）
    const dateTimeStr = `${year}-${month}-${day} ${hourStr}:${minuteStr}:${secondStr}`;
    
    // 使用 Utilities.parseDate 来解析指定时区的日期时间字符串
    // 这会返回一个 Date 对象，表示该时区的本地时间对应的 UTC 时间
    const finalDate = Utilities.parseDate(dateTimeStr, tz, 'yyyy-MM-dd HH:mm:ss');
    
    Logger.log(`解析日期时间: ${dateInput} ${timeInput} (时区: ${tz}) -> ${finalDate.toISOString()}`);
    
    return finalDate;
  } catch (error) {
    Logger.log(`日期时间解析错误: ${dateInput} (${typeof dateInput}) ${timeInput} (${typeof timeInput}) - ${error.message}`);
    return null;
  }
}

/**
 * 获取时区偏移量（相对于 UTC，单位：分钟）
 * @param {string} timezone - 时区标识符（如 Asia/Shanghai）
 * @returns {number} 时区偏移量（分钟）
 */
function getTimezoneOffset(timezone) {
  try {
    const now = new Date();
    // 使用 Utilities.formatDate 来获取指定时区的当前时间
    const localTimeStr = Utilities.formatDate(now, timezone, 'yyyy-MM-dd HH:mm:ss');
    const utcTimeStr = Utilities.formatDate(now, 'UTC', 'yyyy-MM-dd HH:mm:ss');
    
    // 解析时间字符串并计算差值
    const localTime = new Date(localTimeStr.replace(' ', 'T'));
    const utcTime = new Date(utcTimeStr.replace(' ', 'T'));
    
    // 计算偏移量（分钟）
    const offset = (localTime.getTime() - utcTime.getTime()) / 60000;
    
    return offset;
  } catch (error) {
    Logger.log(`获取时区偏移量失败: ${timezone} - ${error.message}`);
    // 如果失败，返回 0（UTC）
    return 0;
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
  const statusSheetName = CONFIG.STATUS_SHEET_PREFIX + CONFIG.MAIN_SHEET_NAME;
  ensureStatusSheet(spreadsheet, statusSheetName);
  
  const mainSheet = spreadsheet.getSheetByName(CONFIG.MAIN_SHEET_NAME);
  const courses = readCourseData(mainSheet);
  
  if (courses.length > 0) {
    const statusSheet = spreadsheet.getSheetByName(statusSheetName);
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


