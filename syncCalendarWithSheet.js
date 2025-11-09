/**
 * Google Apps Script: 同步课程信息到日历
 * 
 * 功能：
 * 1. 从Google表格读取课程信息
 * 2. 在组织者日历上创建事件，老师和学生作为受邀者
 * 3. 系统自动发送邀请邮件给老师和学生（通过 Google Calendar 的邀请功能，无需主动发送）
 * 4. 在隐藏sheet中记录处理状态
 * 
 * 注意：
 * - 创建事件时使用 sendInvites: true，Google Calendar 会自动发送邀请邮件给所有受邀者
 * - 不需要主动发送邮件给老师和学生，系统会自动处理
 * - 只有在取消课程时才会主动发送取消邮件
 * 
 * 架构设计：
 * - 配置表（_SheetConfig）：管理要处理的Sheet列表和配置信息
 * - 状态表（_StatusLog_{SheetName}）：记录每条课程的处理状态和事件ID
 * - 主课程表：包含课程信息（课次、日期、课程内容、时间、老师、学生）
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
    Logger.log('菜单执行同步：开始');
    const ui = SpreadsheetApp.getUi();
    
    Logger.log('菜单执行同步：显示确认对话框');
    const response = ui.alert(
      '确认执行同步',
      '这将处理所有配置的课程表，在组织者日历上创建事件并邀请老师和学生。\n\n是否继续？',
      ui.ButtonSet.YES_NO
    );
    
    Logger.log('菜单执行同步：用户响应 = ' + response);
    
    if (response === ui.Button.YES) {
      Logger.log('菜单执行同步：用户确认，开始执行 main()');
      
      try {
        // 执行主函数
        main();
        
        Logger.log('菜单执行同步：main() 执行完成，显示完成提示');
        // 显示完成提示
        ui.alert(
          '同步完成',
          '课程同步已完成，请查看执行日志了解详细信息。',
          ui.ButtonSet.OK
        );
      } catch (mainError) {
        Logger.log('菜单执行同步：main() 执行失败: ' + mainError.message);
        if (mainError.stack) {
          Logger.log('菜单执行同步：main() 错误堆栈: ' + mainError.stack);
        }
        throw mainError; // 重新抛出，让外层 catch 处理
      }
    } else {
      Logger.log('菜单执行同步：用户取消');
    }
  } catch (error) {
    Logger.log('菜单执行同步：捕获到错误');
    Logger.log('错误类型: ' + (error.name || 'Unknown'));
    Logger.log('错误消息: ' + (error.message || error.toString() || '未知错误'));
    if (error.stack) {
      Logger.log('错误堆栈: ' + error.stack);
    }
    
    try {
      const ui = SpreadsheetApp.getUi();
      const errorMessage = error.message || error.toString() || '未知错误';
      const errorStack = error.stack ? '\n\n错误堆栈:\n' + error.stack.substring(0, 500) : ''; // 限制堆栈长度
      ui.alert(
        '执行错误',
        '同步过程中发生错误：\n' + errorMessage + errorStack + '\n\n请查看执行日志了解详细信息。',
        ui.ButtonSet.OK
      );
    } catch (uiError) {
      // 如果 UI 操作也失败，至少记录到日志
      Logger.log('无法显示错误对话框: ' + uiError.message);
    }
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
      <p><strong>版本：</strong>3.0（组织者模式）</p>
      <p><strong>功能：</strong></p>
      <ul>
        <li>从配置表读取多个课程表</li>
        <li>在组织者日历上创建事件</li>
        <li>自动邀请老师和学生（作为受邀者）</li>
        <li>系统自动发送邀请邮件</li>
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

// ==================== 工具函数 ====================

/**
 * 清理表头文本，去除格式和不可见字符
 * @param {string} text - 原始文本
 * @returns {string} 清理后的文本
 */
function cleanHeaderText(text) {
  if (!text) return '';
  // 转换为字符串
  let cleaned = String(text);
  // 去除所有空白字符（包括空格、制表符、换行符等）
  cleaned = cleaned.replace(/\s+/g, '');
  // 去除不可见字符（零宽字符等）
  cleaned = cleaned.replace(/[\u200B-\u200D\uFEFF]/g, '');
  // 转换为小写
  cleaned = cleaned.toLowerCase();
  return cleaned;
}

// ==================== 主函数 ====================

/**
 * 主执行函数 - 处理所有课程记录
 * 从配置表 _SheetConfig 读取要处理的 sheet 列表，然后循环处理每个 sheet
 */
function main() {
  try {
    Logger.log('通知\t已开始执行');
    Logger.log('main() 函数开始执行');
    
    Logger.log('获取当前表格对象');
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      throw new Error('无法获取当前表格对象，请确保在 Google 表格中运行此脚本');
    }
    Logger.log('表格对象获取成功: ' + spreadsheet.getName());
    
    // 从配置表读取要处理的 sheet 配置信息
    Logger.log('开始读取配置表');
    const sheetConfigMap = readSheetConfig(spreadsheet);
    Logger.log('配置表读取完成');
    
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
    let totalRecordsSuccess = 0;
    let totalRecordsFailed = 0;
    for (const result of allResults) {
      // 判断 Sheet 是否成功：如果没有错误且没有失败的记录，则算作成功
      const sheetSuccess = result.success && result.failed === 0;
      if (sheetSuccess) {
        totalSuccess++;
      } else {
        totalFailed++;
      }
      totalProcessed += result.processed;
      totalRecordsSuccess += (result.processed - result.failed);
      totalRecordsFailed += result.failed;
      
      const status = sheetSuccess ? '成功' : '失败';
      Logger.log(`${result.sheetName}: ${status} - 处理 ${result.processed} 条记录`);
    }
    
    Logger.log(`\n=== 所有 Sheet 处理结果汇总 ===`);
    Logger.log(`总计: 成功 ${totalRecordsSuccess}, 失败 ${totalRecordsFailed}, 共处理 ${totalProcessed} 条记录`);
    
    Logger.log('通知\t执行完毕');
    
  } catch (error) {
    const errorMessage = error.message || error.toString() || '未知错误';
    Logger.log(`主函数执行失败: ${errorMessage}`);
    if (error.stack) {
      Logger.log(`错误堆栈: ${error.stack}`);
    }
    // 记录更详细的错误信息
    Logger.log(`错误类型: ${error.name || 'Unknown'}`);
    Logger.log(`错误详情: ${JSON.stringify(error, null, 2)}`);
    throw error;
  }
}

/**
 * 从配置表读取要处理的 Sheet 配置信息
 * @param {Spreadsheet} spreadsheet - 表格对象
 * @returns {Map<string, Object>} Sheet 配置信息映射表，key为Sheet名称，value为配置对象
 */
function readSheetConfig(spreadsheet) {
  try {
    Logger.log('readSheetConfig: 开始读取配置表');
    
    // 先列出所有 sheet，用于调试
    Logger.log('readSheetConfig: 获取所有 Sheet');
    const allSheets = spreadsheet.getSheets();
    const allSheetNames = allSheets.map(s => s.getName());
    Logger.log(`当前表格中的所有 Sheet: ${allSheetNames.join(', ')}`);
    Logger.log(`正在查找配置表: ${CONFIG.CONFIG_SHEET_NAME}`);
    
    Logger.log('readSheetConfig: 查找配置表 Sheet');
    const configSheet = spreadsheet.getSheetByName(CONFIG.CONFIG_SHEET_NAME);
    
    // 如果配置表不存在，直接报错
    if (!configSheet) {
      const errorMsg = `配置表 ${CONFIG.CONFIG_SHEET_NAME} 不存在，请先创建配置表。当前表格中的 Sheet: ${allSheetNames.join(', ')}`;
      Logger.log('readSheetConfig: 错误 - ' + errorMsg);
      throw new Error(errorMsg);
    }
    
    Logger.log(`✓ 找到配置表: ${CONFIG.CONFIG_SHEET_NAME}`);
    
    // 读取配置表数据
    Logger.log('readSheetConfig: 读取配置表数据');
    const dataRange = configSheet.getDataRange();
    // 使用 getDisplayValues() 获取显示值，避免格式问题
    const values = dataRange.getDisplayValues();
    
    Logger.log(`配置表数据行数: ${values.length}`);
    
    if (values.length < 2) {
      const errorMsg = `配置表 ${CONFIG.CONFIG_SHEET_NAME} 没有数据（只有表头），请至少添加一行数据`;
      Logger.log('readSheetConfig: 错误 - ' + errorMsg);
      throw new Error(errorMsg);
    }
    
    // 解析表头 - 清理格式和不可见字符
    Logger.log('readSheetConfig: 解析表头');
    const headers = values[0];
    Logger.log(`配置表表头（原始）: ${headers.join(', ')}`);
    
    const headerMap = {};
    headers.forEach((header, index) => {
      // 先获取原始值
      const rawHeader = String(header || '').trim();
      // 清理后的表头（用于匹配）
      const normalizedHeader = cleanHeaderText(rawHeader);
      headerMap[normalizedHeader] = index;
      // 同时存储原始表头（用于调试）
      Logger.log(`  表头[${index}]: "${rawHeader}" -> 清理后: "${normalizedHeader}"`);
    });
    Logger.log('readSheetConfig: 表头映射完成');
    Logger.log('headerMap 键: ' + Object.keys(headerMap).join(', '));
    
    // 支持多种表头名称（更宽松的匹配）
    // 先尝试精确匹配（使用清理后的文本）
    Logger.log('开始匹配 Sheet名称 列...');
    
    // 定义可能的匹配键（清理后的格式）
    const possibleKeys = [
      cleanHeaderText('Sheet名称'),
      cleanHeaderText('Sheet Name'),
      cleanHeaderText('名称'),
      cleanHeaderText('Name'),
      cleanHeaderText('Sheet'),
      cleanHeaderText('表名')
    ];
    
    let sheetNameHeader = undefined;
    for (const key of possibleKeys) {
      Logger.log(`尝试匹配: "${key}"`);
      if (headerMap[key] !== undefined) {
        sheetNameHeader = headerMap[key];
        Logger.log(`✓ 找到匹配: "${key}" (索引: ${sheetNameHeader})`);
        break;
      }
    }
    
    // 如果精确匹配失败，尝试模糊匹配（包含关键词）
    if (sheetNameHeader === undefined) {
      Logger.log('精确匹配失败，尝试模糊匹配...');
      for (const [key, index] of Object.entries(headerMap)) {
        // 检查是否包含关键词
        if (key.includes('sheet') && (key.includes('名称') || key.includes('name'))) {
          sheetNameHeader = index;
          Logger.log(`找到匹配的表头: "${headers[index]}" (索引: ${index}, 键: "${key}")`);
          break;
        }
        if (key === '名称' || key === 'name' || key === 'sheet' || key === '表名') {
          sheetNameHeader = index;
          Logger.log(`找到匹配的表头: "${headers[index]}" (索引: ${index}, 键: "${key}")`);
          break;
        }
      }
    }
    
    if (sheetNameHeader !== undefined) {
      Logger.log(`✓ Sheet名称 列匹配成功: 索引 ${sheetNameHeader}, 表头: "${headers[sheetNameHeader]}"`);
    } else {
      Logger.log('✗ Sheet名称 列匹配失败');
    }
    
    // 辅助函数：使用清理后的文本匹配表头
    function findHeaderIndex(possibleNames) {
      for (const name of possibleNames) {
        const cleanedName = cleanHeaderText(name);
        if (headerMap[cleanedName] !== undefined) {
          return headerMap[cleanedName];
        }
      }
      return undefined;
    }
    
    const enabledHeader = findHeaderIndex([
      '启用状态', 'enabled', '启用', '状态', 'status', '是否启用', 'enable', 'active'
    ]);
    
    // 组织者日历ID（必需）
    const organizerCalendarIdHeader = findHeaderIndex([
      '组织者日历ID', '组织者日历id', 'organizer calendar id', '组织者日历', 
      'organizer calendar', '组织者日历授权ID', '组织者日历授权id', 
      '管理员日历ID', 'admin calendar id', '管理员日历', 'admin calendar'
    ]);
    
    const teacherEmailHeader = findHeaderIndex([
      '老师邮箱', 'teacher email', '老师email', 'teacheremail', '老师邮件'
    ]);
    
    const studentEmailHeader = findHeaderIndex([
      '学生邮箱', 'student email', '学生email', 'studentemail', '学生邮件'
    ]);
    
    const timezoneHeader = findHeaderIndex([
      '时区', 'timezone', 'time zone', 'tz'
    ]);
    
    const reminderMinutesHeader = findHeaderIndex([
      '提醒时间', 'reminder minutes', 'reminder', '提醒', 
      '邮件提醒', 'email reminder', '提前提醒', 'minutes before'
    ]);
    
    // 检查必需字段
    if (sheetNameHeader === undefined) {
      // 最后尝试：直接遍历 headerMap 查找包含关键词的键
      Logger.log('最后尝试：遍历 headerMap 查找包含关键词的键...');
      for (const [key, index] of Object.entries(headerMap)) {
        Logger.log(`  检查键: "${key}" (索引: ${index})`);
        if (key.includes('sheet') && (key.includes('名称') || key.includes('name'))) {
          sheetNameHeader = index;
          Logger.log(`  找到匹配的键: "${key}" (索引: ${index})`);
          break;
        }
      }
    }
    
    if (sheetNameHeader === undefined) {
      const availableHeaders = Object.keys(headerMap).join(', ');
      const errorMsg = `配置表 ${CONFIG.CONFIG_SHEET_NAME} 缺少"Sheet名称"列。\n当前表头: ${headers.join(', ')}\n可用的表头键: ${availableHeaders}\n请确保包含 Sheet 名称的列，支持的列名：Sheet名称、Sheet Name、名称、Name、Sheet、表名等`;
      Logger.log('错误: ' + errorMsg);
      throw new Error(errorMsg);
    }
    
    if (organizerCalendarIdHeader === undefined) {
      throw new Error(`配置表 ${CONFIG.CONFIG_SHEET_NAME} 缺少"组织者日历ID"列。\n当前表头: ${headers.join(', ')}\n请确保包含组织者日历ID的列，支持的列名：组织者日历ID、Organizer Calendar ID、组织者日历、管理员日历ID等`);
    }
    
    // 读取启用的 Sheet 配置信息
    const sheetConfigMap = new Map();
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const sheetName = row[sheetNameHeader];
      
      // 跳过空行
      if (!sheetName || String(sheetName).trim() === '') {
        continue;
      }
      
      const sheetNameTrimmed = String(sheetName).trim();
      
      // 检查启用状态
      if (enabledHeader !== undefined) {
        const enabled = row[enabledHeader];
        const enabledStr = String(enabled).trim().toLowerCase();
        // 支持多种表示方式：是/Yes/1/true/启用
        if (enabledStr !== '是' && enabledStr !== 'yes' && enabledStr !== '1' && enabledStr !== 'true' && enabledStr !== '启用' && enabledStr !== 'enabled') {
          Logger.log(`跳过未启用的 Sheet: ${sheetNameTrimmed}`);
          continue;
        }
      }
      
      // 验证 Sheet 是否存在
      const sheet = spreadsheet.getSheetByName(sheetNameTrimmed);
      if (!sheet) {
        Logger.log(`警告：配置的 Sheet "${sheetNameTrimmed}" 不存在，已跳过`);
        continue;
      }
      
      // 读取组织者日历ID（必需）
      const organizerCalendarId = row[organizerCalendarIdHeader] ? String(row[organizerCalendarIdHeader]).trim() : '';
      if (!organizerCalendarId) {
        Logger.log(`警告：组织者日历ID为空，跳过 Sheet: ${sheetNameTrimmed}`);
        continue;
      }
      
      // 读取提醒时间
      let reminderMinutesStr = '';
      if (reminderMinutesHeader !== undefined && row[reminderMinutesHeader] !== undefined && row[reminderMinutesHeader] !== null && row[reminderMinutesHeader] !== '') {
        reminderMinutesStr = String(row[reminderMinutesHeader]).trim();
      }
      
      let reminderMinutes = null;
      if (reminderMinutesStr) {
        const parsed = parseInt(reminderMinutesStr, 10);
        if (!isNaN(parsed) && parsed > 0) {
          reminderMinutes = parsed;
        }
      }
      
      const config = {
        sheetName: sheetNameTrimmed,
        organizerCalendarId: organizerCalendarId,
        teacherEmail: teacherEmailHeader !== undefined ? (row[teacherEmailHeader] || '').trim() : '',
        studentEmail: studentEmailHeader !== undefined ? (row[studentEmailHeader] || '').trim() : '',
        timezone: timezoneHeader !== undefined ? (row[timezoneHeader] || '').trim() : CONFIG.TIMEZONE,
        reminderMinutes: reminderMinutes
      };
      
      // 如果时区为空，使用默认时区
      if (!config.timezone) {
        config.timezone = CONFIG.TIMEZONE;
      }
      
      Logger.log(`  ✓ 添加 Sheet: ${sheetNameTrimmed}`);
      Logger.log(`    组织者日历ID: ${config.organizerCalendarId}`);
      Logger.log(`    老师邮箱: ${config.teacherEmail}`);
      Logger.log(`    学生邮箱: ${config.studentEmail}`);
      Logger.log(`    时区: ${config.timezone}`);
      Logger.log(`    提醒时间: ${config.reminderMinutes ? config.reminderMinutes + '分钟' : '未配置'}`);
      
      sheetConfigMap.set(sheetNameTrimmed, config);
    }
  
    Logger.log(`从配置表读取到 ${sheetConfigMap.size} 个启用的 Sheet 配置`);
    Logger.log('readSheetConfig: 配置读取完成');
    return sheetConfigMap;
    
  } catch (error) {
    Logger.log('readSheetConfig: 捕获到错误');
    Logger.log('错误类型: ' + (error.name || 'Unknown'));
    Logger.log('错误消息: ' + (error.message || error.toString() || '未知错误'));
    if (error.stack) {
      Logger.log('错误堆栈: ' + error.stack);
    }
    throw error;
  }
}

// ==================== 第三部分：课程数据处理和状态管理 ====================

/**
 * 处理单个 Sheet
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
          cancelCourse(deletedRecord, statusSheet, config);
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
      const hasOrganizerEventId = existingRecord.organizerEventId && String(existingRecord.organizerEventId).trim() !== '';
      
      if (hasOrganizerEventId) {
        // 验证事件是否真实存在于日历中
        let organizerEventExists = false;
        let needRecreate = false;
        
        // 验证组织者日历事件（只有当事件ID非空时才验证）
        if (hasOrganizerEventId && existingRecord.organizerCalendarId) {
          try {
            organizerEventExists = verifyCalendarEventExists(existingRecord.organizerCalendarId, existingRecord.organizerEventId);
            if (!organizerEventExists) {
              Logger.log(`[${sheetName}] 组织者日历事件不存在（可能被删除）: ${existingRecord.organizerEventId}，将重新创建`);
              needRecreate = true;
              // 更新状态表，清除无效的事件ID
              statusSheet.getRange(existingRecord.rowIndex, 6).setValue(''); // 第6列是组织者日历事件ID
              existingRecord.organizerEventId = '';
            }
          } catch (error) {
            Logger.log(`[${sheetName}] 验证组织者日历事件失败: ${existingRecord.organizerEventId} - ${error.message}`);
            organizerEventExists = false; // 验证失败，认为不存在
            needRecreate = true;
            // 更新状态表，清除无效的事件ID
            statusSheet.getRange(existingRecord.rowIndex, 6).setValue('');
            existingRecord.organizerEventId = '';
          }
        } else if (hasOrganizerEventId) {
          // 有事件ID但没有日历ID，无法验证，需要重新创建
          Logger.log(`[${sheetName}] 组织者日历事件ID存在但缺少日历ID，将重新创建`);
          needRecreate = true;
          statusSheet.getRange(existingRecord.rowIndex, 6).setValue('');
          existingRecord.organizerEventId = '';
        }
        
        // 如果事件存在，跳过处理
        if (organizerEventExists) {
          Logger.log(`[${sheetName}] 跳过处理（token相同且日历事件已验证存在）: ${course.lessonNumber}`);
          return false;
        }
        
        // 如果有事件不存在或需要重新创建，需要重新处理
        if (needRecreate || !organizerEventExists) {
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
    for (let i = 0; i < toProcess.length; i++) {
      const course = toProcess[i];
      try {
        const result = processCourse(course, statusSheet, config);
        results.push(result);
        Logger.log(`[${sheetName}] 处理完成: ${course.lessonNumber} - ${result.status}`);
        
        // 如果不是最后一条记录，添加延迟，避免连续处理多条记录时触发速率限制
        if (i < toProcess.length - 1) {
          addOperationDelay();
        }
      } catch (error) {
        Logger.log(`[${sheetName}] 处理失败: ${course.lessonNumber} - ${error.message}`);
        results.push({
          course: course,
          status: '失败',
          error: error.message
        });
        
        // 即使失败，也添加延迟，避免连续处理时触发速率限制
        if (i < toProcess.length - 1) {
          addOperationDelay();
        }
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
 * 读取课程数据
 */
function readCourseData(sheet, config) {
  const dataRange = sheet.getDataRange();
  // 使用 getDisplayValues() 获取显示值，避免格式问题
  const values = dataRange.getDisplayValues();
  
  if (values.length < 2) {
    return [];
  }
  
  // 表头行（第1行，索引0）
  const headers = values[0];
  const headerMap = {};
  headers.forEach((header, index) => {
    // 使用清理后的文本作为键，但保留原始表头用于匹配
    const rawHeader = String(header || '').trim();
    const cleanedHeader = cleanHeaderText(rawHeader);
    // 同时存储原始表头和清理后的表头
    headerMap[rawHeader] = index;
    headerMap[cleanedHeader] = index;
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
        organizerCalendarId: config.organizerCalendarId || '',
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
      if (!course.date || !course.organizerCalendarId) {
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
  const organizerCalendarIdCol = getColumnIndex(['组织者日历id', 'organizer calendar id', '组织者日历', 'organizer calendar', '管理员日历id', 'admin calendar id']);
  const organizerEventIdCol = getColumnIndex(['组织者日历事件id', 'organizer event id', '组织者事件id', 'organizer event id', '管理员日历事件id', 'admin event id']);
  const organizerEventTimeCol = getColumnIndex(['组织者日历创建时间', 'organizer event time', '组织者事件时间', 'organizer event time', '管理员日历创建时间', 'admin event time']);
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
    
    // 读取组织者日历ID和事件ID（确保不是Date对象）
    let organizerCalendarId = getValue(organizerCalendarIdCol);
    if (organizerCalendarId instanceof Date) {
      organizerCalendarId = '';
    } else {
      organizerCalendarId = String(organizerCalendarId).trim();
    }
    
    let organizerEventId = getValue(organizerEventIdCol);
    if (organizerEventId instanceof Date) {
      organizerEventId = '';
    } else {
      organizerEventId = String(organizerEventId).trim();
    }
    
    const record = {
      recordId: recordId, // 记录ID
      lessonNumber: lessonNumber,
      date: date,
      token: getValue(tokenCol), // Token（关键信息哈希）
      organizerCalendarId: organizerCalendarId, // 组织者日历ID（用于删除事件）
      organizerEventId: organizerEventId, // 组织者日历事件ID
      status: getValue(statusCol), // 处理状态
      rowIndex: i + 1 // 状态表的行号（从1开始，包含表头）
    };
    
    // 验证事件ID格式：如果事件ID是"已发送"或其他状态文本，说明是错误的数据，应该清空
    const invalidStatusTexts = ['已发送', '未发送', '失败', '部分失败', '已完成', '处理中'];
    if (record.organizerEventId && invalidStatusTexts.includes(record.organizerEventId)) {
      Logger.log(`警告：组织者事件ID包含状态文本，将被清空: "${record.organizerEventId}"`);
      record.organizerEventId = '';
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
    const statusRow = statusSheet.getRange(course.rowIndex, 1, 1, statusSheet.getLastColumn()).getValues()[0];
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
    organizerEventId: existingRecord ? (existingRecord.organizerEventId || null) : null,
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
        organizerCalendarId: record.organizerCalendarId || '',
        organizerEventId: record.organizerEventId || '',
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
        organizerCalendarId: record.organizerCalendarId || '',
        organizerEventId: record.organizerEventId || '',
        rowIndex: record.rowIndex,
        token: record.token || ''
      });
    }
  });
  
  return deletedRecords;
}

/**
 * 查找相同课次但不同日期的旧记录（日期变化）
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
  const organizerCalendarIdCol = getColumnIndex(['组织者日历id', 'organizer calendar id', '组织者日历', 'organizer calendar', '管理员日历id', 'admin calendar id']);
  const organizerEventIdCol = getColumnIndex(['组织者日历事件id', 'organizer event id', '组织者事件id', 'organizer event id', '管理员日历事件id', 'admin event id']);
  const recordIdCol = getColumnIndex(['记录id', 'record id', '记录id', 'recordid', 'id']);
  
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
        // 获取记录ID（如果存在）
        const recordId = getValue(row, recordIdCol);
        
        oldRecords.push({
          recordId: recordId, // 添加记录ID，用于判断是否是同一条记录
          lessonNumber: rowLessonNumber,
          date: rowDate,
          organizerCalendarId: getValue(row, organizerCalendarIdCol),
          organizerEventId: getValue(row, organizerEventIdCol),
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

// ==================== 第四部分：日历事件创建和更新（组织者模式） ====================

/**
 * 处理单条课程记录
 */
function processCourse(course, statusSheet, config) {
  const result = {
    course: course,
    organizerEvent: { eventId: null, error: null },
    status: '处理中'
  };
  
  try {
    // 获取已有的事件ID和token信息（在删除旧记录之前获取，以便判断是否应该更新）
    const existingInfo = getExistingEventIds(statusSheet, course);
    
    // 如果有旧记录（日期变化），优先尝试更新现有事件，而不是删除后重新创建
    if (course._oldRecords && course._oldRecords.length > 0) {
      // 检查是否有相同记录ID的旧记录（说明是同一条记录，只是日期变化了）
      const sameRecordIdOldRecord = course._oldRecords.find(oldRecord => 
        oldRecord.recordId && course.recordId && oldRecord.recordId === course.recordId
      );
      
      if (sameRecordIdOldRecord) {
        // 如果是同一条记录（记录ID相同），说明只是日期变化，应该更新现有事件而不是删除后重新创建
        Logger.log(`检测到同一条记录的日期变化（记录ID: ${course.recordId}），将更新现有事件而不是删除后重新创建`);
        
        // 将旧记录的事件ID传递给existingInfo，以便后续更新时使用
        if (sameRecordIdOldRecord.organizerEventId && !existingInfo.organizerEventId) {
          existingInfo.organizerEventId = sameRecordIdOldRecord.organizerEventId;
          Logger.log(`使用旧记录的组织者事件ID进行更新: ${sameRecordIdOldRecord.organizerEventId}`);
        }
        
        // 删除其他不同记录ID的旧记录（这些是真正的旧记录，需要删除）
        const otherOldRecords = course._oldRecords.filter(oldRecord => 
          !oldRecord.recordId || oldRecord.recordId !== course.recordId
        );
        
        if (otherOldRecords.length > 0) {
          Logger.log(`删除 ${otherOldRecords.length} 条其他旧记录`);
          for (const oldRecord of otherOldRecords) {
            // 尝试删除组织者日历事件
            if (oldRecord.organizerEventId) {
              try {
                if (oldRecord.organizerCalendarId) {
                  deleteCalendarEvent(oldRecord.organizerCalendarId, oldRecord.organizerEventId);
                  Logger.log(`删除旧组织者日历事件成功: ${oldRecord.organizerEventId}`);
                } else {
                  deleteCalendarEventById(oldRecord.organizerEventId);
                  Logger.log(`删除旧组织者日历事件成功: ${oldRecord.organizerEventId}`);
                }
                addOperationDelay();
              } catch (error) {
                Logger.log(`删除旧组织者日历事件失败: ${oldRecord.organizerEventId} - ${error.message}`);
              }
            }
          }
          
          // 删除其他旧记录的状态记录
          deleteOldStatusRecords(statusSheet, otherOldRecords);
        }
      } else {
        // 如果没有相同记录ID的旧记录，说明是真正的旧记录，需要删除
        Logger.log(`检测到 ${course._oldRecords.length} 条旧记录，将删除这些旧记录`);
        for (const oldRecord of course._oldRecords) {
          // 尝试删除组织者日历事件
          if (oldRecord.organizerEventId) {
            try {
              if (oldRecord.organizerCalendarId) {
                deleteCalendarEvent(oldRecord.organizerCalendarId, oldRecord.organizerEventId);
                Logger.log(`删除旧组织者日历事件成功: ${oldRecord.organizerEventId}`);
              } else {
                deleteCalendarEventById(oldRecord.organizerEventId);
                Logger.log(`删除旧组织者日历事件成功: ${oldRecord.organizerEventId}`);
              }
              addOperationDelay();
            } catch (error) {
              Logger.log(`删除旧组织者日历事件失败: ${oldRecord.organizerEventId} - ${error.message}`);
            }
          }
        }
        
        // 删除旧状态记录
        deleteOldStatusRecords(statusSheet, course._oldRecords);
      }
    }
    
    // 判断是否需要更新事件（关键信息有变化时）
    const needsUpdate = existingInfo.hasChanges;
    
    // 创建或更新组织者日历事件（在组织者日历上创建，老师和学生作为受邀者）
    // 系统会自动发送邀请邮件给受邀者
    if (needsUpdate || !existingInfo.organizerEventId) {
      try {
        // 在组织者日历上创建或更新事件，添加老师和学生作为受邀者
        const organizerEventId = createOrUpdateCalendarEvent(
          config.organizerCalendarId,
          course,
          existingInfo.organizerEventId,
          config
        );
        if (organizerEventId) {
          result.organizerEvent.eventId = String(organizerEventId);
          if (existingInfo.organizerEventId && needsUpdate) {
            Logger.log(`组织者日历事件更新成功: ${organizerEventId}，已通知所有受邀人`);
          } else if (existingInfo.organizerEventId) {
            Logger.log(`组织者日历事件保持不变: ${organizerEventId}`);
          } else {
            Logger.log(`组织者日历事件创建成功: ${organizerEventId}，已邀请老师和学生`);
          }
        } else {
          result.organizerEvent.error = '创建事件成功但未返回事件ID';
          Logger.log(`组织者日历事件处理失败: 创建事件成功但未返回事件ID`);
        }
        // 添加延迟，避免速率限制
        addOperationDelay();
      } catch (error) {
        result.organizerEvent.error = error.message;
        Logger.log(`组织者日历事件处理失败: ${error.message}`);
        // 如果是速率限制错误，记录详细信息
        if (isRateLimitError(error)) {
          Logger.log(`⚠️ 组织者日历事件遇到速率限制，可能需要稍后重试`);
        }
        // 即使创建失败，也尝试保留已有的事件ID（如果有）
        if (existingInfo.organizerEventId) {
          result.organizerEvent.eventId = String(existingInfo.organizerEventId);
          Logger.log(`保留已有组织者日历事件ID: ${existingInfo.organizerEventId}`);
        }
      }
    } else {
      // token相同且已有事件ID，跳过更新
      result.organizerEvent.eventId = existingInfo.organizerEventId ? String(existingInfo.organizerEventId) : null;
      Logger.log(`组织者日历事件跳过（token相同且已有事件）: ${existingInfo.organizerEventId}`);
    }
    
    // 判断整体状态
    const organizerEventId = result.organizerEvent.eventId ? String(result.organizerEvent.eventId).trim() : '';
    const organizerSuccess = organizerEventId !== '' && !result.organizerEvent.error;
    
    Logger.log(`[${course.lessonNumber}] 状态判断: 组织者事件ID=${organizerEventId || '无'}, 成功=${organizerSuccess}`);
    
    if (organizerSuccess) {
      result.status = '已完成';
    } else {
      result.status = '失败';
    }
    
    Logger.log(`[${course.lessonNumber}] 最终状态: ${result.status}`);
    
    // 记录状态到隐藏sheet
    updateStatusRecord(statusSheet, course, result);
    
    return result;
    
  } catch (error) {
    result.status = '失败';
    result.error = error.message;
    updateStatusRecord(statusSheet, course, result);
    throw error;
  }
}

// ==================== 第五部分：删除和取消功能 ====================

/**
 * 取消课程（删除日历事件并发送取消邮件）
 */
function cancelCourse(deletedRecord, statusSheet, config) {
  // 从状态表中获取日历ID和事件ID信息
  // deletedRecord 已经包含了 organizerEventId
  // 还需要获取日历ID（组织者日历ID）
  
  // 读取状态表中的完整信息（作为备用）
  const headerRow = statusSheet.getRange(1, 1, 1, statusSheet.getLastColumn()).getValues()[0];
  const headerMap = {};
  headerRow.forEach((header, index) => {
    const headerKey = String(header).trim().toLowerCase();
    headerMap[headerKey] = index;
  });
  
  const getColumnIndex = (headerNames) => {
    for (const name of headerNames) {
      const key = name.toLowerCase();
      if (headerMap[key] !== undefined) {
        return headerMap[key];
      }
    }
    return undefined;
  };
  
  const organizerCalendarIdCol = getColumnIndex(['组织者日历id', 'organizer calendar id', '组织者日历', 'organizer calendar', '管理员日历id', 'admin calendar id']);
  const organizerEventIdCol = getColumnIndex(['组织者日历事件id', 'organizer event id', '组织者事件id', 'organizer event id', '管理员日历事件id', 'admin event id']);
  
  const statusRow = statusSheet.getRange(deletedRecord.rowIndex, 1, 1, statusSheet.getLastColumn()).getValues()[0];
  
  // 获取日历ID（优先使用deletedRecord中的，如果为空则从状态表中读取，最后使用config中的）
  const organizerCalendarId = deletedRecord.organizerCalendarId || 
                              (organizerCalendarIdCol !== undefined ? statusRow[organizerCalendarIdCol] : '') || 
                              (config ? config.organizerCalendarId : '') || '';
  const organizerEventId = deletedRecord.organizerEventId || 
                           (organizerEventIdCol !== undefined ? statusRow[organizerEventIdCol] : '') || '';
  
  // 1. 删除组织者日历事件
  if (organizerEventId) {
    try {
      if (organizerCalendarId) {
        // 如果有日历ID，直接删除
        deleteCalendarEvent(organizerCalendarId, organizerEventId);
        Logger.log(`删除组织者日历事件成功: ${organizerEventId} (日历: ${organizerCalendarId})`);
      } else {
        // 如果没有日历ID，尝试通过事件ID删除（遍历所有日历）
        deleteCalendarEventById(organizerEventId);
        Logger.log(`删除组织者日历事件成功: ${organizerEventId}`);
      }
      // 添加延迟，避免速率限制
      addOperationDelay();
    } catch (error) {
      Logger.log(`删除组织者日历事件失败: ${organizerEventId} - ${error.message}`);
      // 如果是速率限制错误，记录详细信息
      if (isRateLimitError(error)) {
        Logger.log(`⚠️ 删除组织者日历事件遇到速率限制，可能需要稍后重试`);
      }
    }
  }
  
  // 2. 发送取消邮件给所有受邀者（老师和学生）
  // 从日历事件中获取参与者信息，或者从config中获取
  try {
    sendCancellationEmails(deletedRecord, config);
  } catch (error) {
    Logger.log(`发送取消邮件失败: ${error.message}`);
  }
  
  // 3. 清空状态记录（保留行，但清空内容）
  const emptyRow = ['', '', '', '', '', '', '', '', '']; // 9列（包含记录ID和组织者日历ID）
  statusSheet.getRange(deletedRecord.rowIndex, 1, 1, emptyRow.length).setValues([emptyRow]);
}

/**
 * 发送课程取消邮件
 */
function sendCancellationEmails(deletedRecord, config) {
  // 由于记录已被删除，我们需要从日历事件中获取参与者信息
  // 或者从状态表中获取之前保存的信息
  
  // 尝试从日历事件中获取参与者信息
  let event = null;
  let calendar = null;
  
  // 尝试通过组织者日历事件ID获取
  if (deletedRecord.organizerEventId) {
    try {
      // 优先使用组织者日历ID
      if (deletedRecord.organizerCalendarId) {
        calendar = getCalendarByIdOrEmail(deletedRecord.organizerCalendarId, null);
        if (calendar) {
          event = calendar.getEventById(deletedRecord.organizerEventId);
        }
      }
      
      // 如果没找到，尝试遍历所有日历
      if (!event) {
        const calendars = CalendarApp.getAllCalendars();
        for (const cal of calendars) {
          try {
            event = cal.getEventById(deletedRecord.organizerEventId);
            if (event) {
              calendar = cal;
              break;
            }
          } catch (error) {
            continue;
          }
        }
      }
    } catch (error) {
      Logger.log(`获取日历事件失败: ${error.message}`);
    }
  }
  
  // 如果无法从事件中获取参与者信息，使用config中的邮箱
  let teacherEmail = null;
  let studentEmail = null;
  
  if (event) {
    // 从事件中获取参与者信息
    const guests = event.getGuestList();
    teacherEmail = guests.length > 0 ? guests[0].getEmail() : null;
    studentEmail = guests.length > 1 ? guests[1].getEmail() : null;
  }
  
  // 如果从事件中无法获取，使用config中的邮箱
  if (!teacherEmail && config && config.teacherEmail) {
    teacherEmail = config.teacherEmail;
  }
  if (!studentEmail && config && config.studentEmail) {
    studentEmail = config.studentEmail;
  }
  
  if (!teacherEmail && !studentEmail) {
    Logger.log(`无法获取参与者邮箱，跳过发送取消邮件`);
    return;
  }
  
  // 构建取消邮件内容
  const courseTitle = event ? (event.getTitle() || '课程') : '课程';
  const eventDate = event ? event.getStartTime() : new Date();
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

// ==================== 第六部分：工具函数和辅助功能 ====================

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
      
      // 如果启用了 Meet 链接，使用 Calendar API 直接创建包含 Meet 链接的事件
      // 这样可以确保 Meet 链接在创建时就存在，所有参与者都能看到
      if (options && options.addMeetLink !== false) {
        try {
          const calendarId = calendar.getId();
          
          // 构建受邀者列表
          const attendees = [];
          if (options && options.guests) {
            const guests = typeof options.guests === 'string' ? 
              options.guests.split(',').map(email => email.trim()).filter(email => email) : 
              options.guests;
            
            for (const guest of guests) {
              if (guest) {
                attendees.push({ email: guest });
              }
            }
          }
          
          // 获取时区（从 course 对象或使用默认时区）
          const timezone = (options && options.timezone) || Session.getScriptTimeZone();
          
          // 格式化日期时间为 RFC3339 格式
          const formatDateTime = (date) => {
            const year = date.getFullYear();
            const month = String(date.getMonth() + 1).padStart(2, '0');
            const day = String(date.getDate()).padStart(2, '0');
            const hours = String(date.getHours()).padStart(2, '0');
            const minutes = String(date.getMinutes()).padStart(2, '0');
            const seconds = String(date.getSeconds()).padStart(2, '0');
            return `${year}-${month}-${day}T${hours}:${minutes}:${seconds}`;
          };
          
          // 使用 Calendar API 创建事件（包含 Meet 链接）
          const eventResource = {
            summary: title,
            description: options && options.description ? options.description : '',
            start: {
              dateTime: formatDateTime(startTime),
              timeZone: timezone
            },
            end: {
              dateTime: formatDateTime(endTime),
              timeZone: timezone
            },
            attendees: attendees,
            conferenceData: {
              createRequest: {
                requestId: Utilities.getUuid(),
                conferenceSolutionKey: {
                  type: 'hangoutsMeet'
                }
              }
            }
          };
          
          // 使用 Calendar API 创建事件
          const createdEvent = Calendar.Events.insert(eventResource, calendarId, {
            sendUpdates: options && options.sendInvites ? 'all' : 'none',
            conferenceDataVersion: 1 // 确保 conferenceData 被处理
          });
          
          // 获取创建的事件对象（用于返回）
          const eventId = createdEvent.id;
          const event = calendar.getEventById(eventId);
          
          Logger.log(`✓ 使用 Calendar API 创建事件（包含 Meet 链接）: ${eventId}`);
          return event;
        } catch (error) {
          // 如果使用 Calendar API 创建失败，回退到使用 CalendarApp
          Logger.log(`⚠️ 使用 Calendar API 创建事件失败，回退到 CalendarApp: ${error.message}`);
          if (error.stack) {
            Logger.log(`错误堆栈: ${error.stack}`);
          }
          // 继续执行，使用 CalendarApp 创建
        }
      }
      
      // 使用 CalendarApp 创建事件（回退方案或未启用 Meet 链接时）
      const event = calendar.createEvent(title, startTime, endTime);
      
      // 设置描述
      if (options && options.description) {
        event.setDescription(options.description);
      }
      
      // 添加受邀者（如果提供了 guests）
      if (options && options.guests) {
        const guests = typeof options.guests === 'string' ? 
          options.guests.split(',').map(email => email.trim()).filter(email => email) : 
          options.guests;
        
        for (const guest of guests) {
          if (guest) {
            event.addGuest(guest);
          }
        }
      }
      
      // 发送邀请（如果设置了 sendInvites）
      if (options && options.sendInvites) {
        // 注意：addGuest 后会自动发送邀请，但我们可以显式设置
        // 实际上，addGuest 已经会自动发送邀请邮件
      }
      
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
        Logger.log(`创建日历事件失败（非速率限制错误）: ${error.message}`);
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
      const newEmails = guests.split(',').map(email => email.trim()).filter(email => email);
      
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

/**
 * 创建或更新日历事件（在组织者日历上创建，老师和学生作为受邀者）
 * @param {string} calendarId - 组织者日历ID
 * @param {Object} course - 课程对象
 * @param {string|null} existingEventId - 已有的事件ID（如果存在则更新，否则创建）
 * @param {Object} config - 配置对象（包含老师和学生邮箱）
 * @returns {string} 事件ID
 */
function createOrUpdateCalendarEvent(calendarId, course, existingEventId, config) {
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
  
  // 构建受邀者列表（老师和学生）
  const guests = [];
  if (config && config.teacherEmail) {
    guests.push(config.teacherEmail);
  }
  if (config && config.studentEmail) {
    guests.push(config.studentEmail);
  }
  // 如果配置中没有邮箱，尝试从课程对象中获取（向后兼容）
  if (guests.length === 0) {
    if (course.teacherEmail) guests.push(course.teacherEmail);
    if (course.studentEmail) guests.push(course.studentEmail);
  }
  const eventGuests = guests.join(',');
  
  let event;
  
  if (existingEventId) {
    // 更新已有事件
    try {
      event = calendar.getEventById(existingEventId);
      
      // 更新事件信息（带速率限制处理）
      updateEventWithRetry(event, eventSummary, eventDescription, eventStart, eventEnd, eventGuests);
      
      // 确保事件有 Google Meet 链接
      try {
        const calendarId = calendar.getId();
        const eventId = existingEventId.split('@')[0]; // 获取事件ID（去掉日历ID后缀）
        
        // 检查事件是否已有 Meet 链接
        const existingEvent = Calendar.Events.get(calendarId, eventId);
        
        // 如果没有 Meet 链接，添加一个
        if (!existingEvent.conferenceData || !existingEvent.conferenceData.entryPoints || 
            existingEvent.conferenceData.entryPoints.length === 0) {
          // Calendar.Events.patch(resource, calendarId, eventId, optionalArgs)
          // 注意：添加 Meet 链接时需要发送更新通知，这样参与者才能看到 Meet 链接
          Calendar.Events.patch({
            conferenceData: {
              createRequest: {
                requestId: Utilities.getUuid(),
                conferenceSolutionKey: {
                  type: 'hangoutsMeet'
                }
              }
            }
          }, calendarId, eventId, {
            sendUpdates: 'all' // 发送更新通知给所有参与者，确保他们能看到 Meet 链接
          });
          
          // 等待一小段时间，让 Meet 链接有时间同步
          Utilities.sleep(500);
          
          Logger.log(`✓ 已为更新的事件添加 Google Meet 链接: ${eventId}`);
        } else {
          Logger.log(`✓ 事件已有 Google Meet 链接: ${eventId}`);
        }
      } catch (error) {
        // 如果添加 Meet 链接失败，记录日志但不影响事件更新
        Logger.log(`⚠️ 添加/检查 Google Meet 链接失败: ${error.message}`);
        if (error.stack) {
          Logger.log(`错误堆栈: ${error.stack}`);
        }
      }
      
      // 更新提醒（如果配置了提醒时间）
      // 注意：提醒会发送给所有参与者，包括组织者和受邀者（老师和学生）
      if (course.reminderMinutes && course.reminderMinutes > 0) {
        try {
          // 清除现有提醒
          event.removeAllReminders();
          // 添加邮件提醒（会发送给所有参与者，包括受邀者）
          event.addEmailReminder(course.reminderMinutes);
          // 添加弹出提醒（在日历应用中显示，适用于所有参与者）
          event.addPopupReminder(course.reminderMinutes);
          Logger.log(`更新提醒: 提前 ${course.reminderMinutes} 分钟（邮件+弹出，所有参与者包括受邀者）`);
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
      sendInvites: true,
      addMeetLink: true, // 添加 Google Meet 链接
      timezone: timezone // 传递时区信息
    }
  );
  
  // 添加提醒（如果配置了提醒时间）
  // 注意：提醒会发送给所有参与者，包括组织者和受邀者（老师和学生）
  if (course.reminderMinutes && course.reminderMinutes > 0) {
    try {
      // 添加邮件提醒（会发送给所有参与者，包括受邀者）
      event.addEmailReminder(course.reminderMinutes);
      // 添加弹出提醒（在日历应用中显示，适用于所有参与者）
      event.addPopupReminder(course.reminderMinutes);
      Logger.log(`添加提醒: 提前 ${course.reminderMinutes} 分钟（邮件+弹出，所有参与者包括受邀者）`);
    } catch (error) {
      Logger.log(`添加提醒失败: ${error.message}`);
      // 提醒失败不影响事件创建，继续执行
    }
  }
  
  Logger.log(`创建新日历事件: ${event.getId()}`);
  return event.getId();
}

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
      '组织者日历ID',      // 4 - 组织者日历ID（用于删除事件）
      '组织者日历事件ID',  // 5 - 组织者日历事件ID
      '组织者日历创建时间',// 6 - 组织者日历创建时间
      '处理状态',          // 7 - 处理状态
      '最后更新时间'       // 8 - 最后更新时间
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
    const emptyRow = ['', '', '', '', '', '', '', '', '']; // 9列（包含记录ID和组织者日历ID）
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
  
  // 获取各列的索引（使用表头名称而不是固定索引）- 适配组织者模式
  const recordIdCol = getColumnIndex(['记录id', 'record id', '记录id', 'recordid', 'id']);
  const lessonNumberCol = getColumnIndex(['课次', 'lesson', 'lesson number', '课程次数']);
  const dateCol = getColumnIndex(['日期', 'date', '课程日期']);
  const tokenCol = getColumnIndex(['token', '令牌', '哈希']);
  const organizerCalendarIdCol = getColumnIndex(['组织者日历id', 'organizer calendar id', '组织者日历', 'organizer calendar', '管理员日历id', 'admin calendar id']);
  const organizerEventIdCol = getColumnIndex(['组织者日历事件id', 'organizer event id', '组织者事件id', 'organizer event id', '管理员日历事件id', 'admin event id']);
  const organizerEventTimeCol = getColumnIndex(['组织者日历创建时间', 'organizer event time', '组织者事件时间', 'organizer event time', '管理员日历创建时间', 'admin event time']);
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
  let existingOrganizerCalendarId = getExistingValue(organizerCalendarIdCol);
  existingOrganizerCalendarId = existingOrganizerCalendarId && !(existingOrganizerCalendarId instanceof Date) ? String(existingOrganizerCalendarId).trim() : '';
  let existingOrganizerEventId = getExistingValue(organizerEventIdCol);
  existingOrganizerEventId = existingOrganizerEventId && !(existingOrganizerEventId instanceof Date) ? String(existingOrganizerEventId).trim() : '';
  
  // 验证事件ID格式：如果事件ID是"已发送"或其他状态文本，说明是错误的数据，应该清空
  const invalidStatusTexts = ['已发送', '未发送', '失败', '部分失败', '已完成', '处理中'];
  if (existingOrganizerEventId && invalidStatusTexts.includes(existingOrganizerEventId)) {
    Logger.log(`警告：组织者事件ID包含状态文本，将被清空: "${existingOrganizerEventId}"`);
    existingOrganizerEventId = '';
  }
  
  // 从config中获取组织者日历ID（如果course中没有）
  const organizerCalendarId = course.organizerCalendarId || existingOrganizerCalendarId || '';
  
  // 确保事件ID是字符串格式，且不是日期对象或状态文本
  let organizerEventId = '';
  if (result.organizerEvent && result.organizerEvent.eventId) {
    const eventId = result.organizerEvent.eventId;
    // 检查是否是日期对象
    if (eventId instanceof Date) {
      Logger.log(`警告：组织者事件ID是日期对象，将被忽略: ${eventId}`);
      organizerEventId = existingOrganizerEventId || '';
    } else {
      const eventIdStr = String(eventId).trim();
      // 验证事件ID格式：如果事件ID是"已发送"或其他状态文本，说明是错误的数据，应该清空
      if (invalidStatusTexts.includes(eventIdStr)) {
        Logger.log(`警告：组织者事件ID包含状态文本，将被忽略: "${eventIdStr}"`);
        organizerEventId = existingOrganizerEventId || '';
      } else {
        organizerEventId = eventIdStr;
      }
    }
  } else {
    organizerEventId = existingOrganizerEventId || '';
  }
  
  // 如果事件ID存在，更新创建时间；如果是新创建的，使用当前时间；如果是已有的，保留原时间
  let organizerEventTime = '';
  
  if (result.organizerEvent && result.organizerEvent.eventId && !(result.organizerEvent.eventId instanceof Date)) {
    // 新创建或更新的事件
    organizerEventTime = nowStr;
  } else if (existingRecord && existingOrganizerEventId) {
    // 保留原有的创建时间
    const existingTime = getExistingValue(organizerEventTimeCol);
    if (existingTime instanceof Date) {
      // 如果是日期对象，格式化为字符串
      const timezone = course.timezone || CONFIG.TIMEZONE || Session.getScriptTimeZone();
      organizerEventTime = Utilities.formatDate(existingTime, timezone, 'yyyy-MM-dd HH:mm:ss');
    } else if (existingTime) {
      organizerEventTime = String(existingTime).trim();
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
    organizerCalendarIdCol, organizerEventIdCol, organizerEventTimeCol,
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
  if (organizerCalendarIdCol !== undefined) rowData[organizerCalendarIdCol] = String(organizerCalendarId || '');
  if (organizerEventIdCol !== undefined) rowData[organizerEventIdCol] = String(organizerEventId || '');
  if (organizerEventTimeCol !== undefined) rowData[organizerEventTimeCol] = String(organizerEventTime || '');
  if (statusCol !== undefined) rowData[statusCol] = result.status;
  if (lastUpdateTimeCol !== undefined) rowData[lastUpdateTimeCol] = nowStr;
  
  // 直接更新对应行（状态表和正式表一一对应）
  statusSheet.getRange(rowIndex, 1, 1, totalCols).setValues([rowData]);
}

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
 * 测试函数 - 用于验证代码是否可以正常运行
 * 在 Google Apps Script 编辑器中运行此函数来测试
 */
function test() {
  try {
    Logger.log('测试开始');
    
    // 测试1: 检查 CONFIG 对象
    Logger.log('测试1: CONFIG 对象');
    Logger.log('CONFIG.CONFIG_SHEET_NAME = ' + CONFIG.CONFIG_SHEET_NAME);
    
    // 测试2: 检查是否可以获取表格对象
    Logger.log('测试2: 获取表格对象');
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      throw new Error('无法获取表格对象');
    }
    Logger.log('表格名称: ' + spreadsheet.getName());
    
    // 测试3: 检查是否可以获取所有 Sheet
    Logger.log('测试3: 获取所有 Sheet');
    const sheets = spreadsheet.getSheets();
    Logger.log('Sheet 数量: ' + sheets.length);
    sheets.forEach((sheet, index) => {
      Logger.log(`  Sheet ${index + 1}: ${sheet.getName()}`);
    });
    
    // 测试4: 检查配置表是否存在
    Logger.log('测试4: 检查配置表');
    const configSheet = spreadsheet.getSheetByName(CONFIG.CONFIG_SHEET_NAME);
    if (configSheet) {
      Logger.log('✓ 配置表存在: ' + CONFIG.CONFIG_SHEET_NAME);
      const dataRange = configSheet.getDataRange();
      const values = dataRange.getValues();
      Logger.log('配置表行数: ' + values.length);
      if (values.length > 0) {
        Logger.log('配置表表头: ' + values[0].join(', '));
      }
    } else {
      Logger.log('✗ 配置表不存在: ' + CONFIG.CONFIG_SHEET_NAME);
    }
    
    Logger.log('测试完成');
    return '测试成功';
    
  } catch (error) {
    Logger.log('测试失败: ' + error.message);
    if (error.stack) {
      Logger.log('错误堆栈: ' + error.stack);
    }
    throw error;
  }
}
