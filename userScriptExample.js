/**
 * 用户表格脚本示例
 * 
 * 使用说明：
 * 1. 在 Google Sheets 中打开新表格
 * 2. 点击"扩展程序" → "Apps Script"
 * 3. 添加库（通过脚本ID）
 * 4. 设置库的标识符（如 CalendarSyncLib）
 * 5. 将以下代码复制到脚本编辑器中
 * 6. 保存并刷新表格
 * 
 * 注意：
 * - 将 CalendarSyncLib 替换为你设置的库标识符
 * - 菜单项函数必须在用户表格中定义，不能直接在库中调用
 */

/**
 * 当打开表格时自动创建自定义菜单
 */
function onOpen() {
  CalendarSyncLib.createMenu();
}

/**
 * 菜单项：执行同步（包装函数）
 * 注意：菜单项函数必须在用户表格中定义，不能直接在库中调用
 */
function menuRunSync() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '确认执行同步',
      '这将处理所有配置的课程表，在组织者日历上创建事件并邀请老师和学生。\n\n是否继续？',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      CalendarSyncLib.main();
      ui.alert(
        '同步完成',
        '课程同步已完成，请查看执行日志了解详细信息。',
        ui.ButtonSet.OK
      );
    }
  } catch (error) {
    const ui = SpreadsheetApp.getUi();
    const errorMessage = error.message || error.toString() || '未知错误';
    ui.alert(
      '执行错误',
      '同步过程中发生错误：\n' + errorMessage + '\n\n请查看执行日志了解详细信息。',
      ui.ButtonSet.OK
    );
  }
}

/**
 * 菜单项：查看配置（包装函数）
 */
function menuViewConfig() {
  CalendarSyncLib.menuViewConfig();
}

/**
 * 菜单项：查看状态表（包装函数）
 */
function menuViewStatus() {
  CalendarSyncLib.menuViewStatus();
}

/**
 * 菜单项：关于（包装函数）
 */
function menuAbout() {
  CalendarSyncLib.menuAbout();
}

/**
 * 可选：手动执行同步
 */
function runSync() {
  CalendarSyncLib.main();
}

