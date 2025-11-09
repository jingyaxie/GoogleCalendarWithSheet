# 库使用说明

本文档说明如何将此脚本作为 Google Apps Script 库使用。

## 发布为库

### 步骤 1：准备代码

代码已经准备好，包含以下关键函数：
- `createMenu()` - 创建自定义菜单
- `main()` - 执行同步（主函数）
- 所有菜单项函数（`menuRunSync`, `menuViewConfig`, `menuViewStatus`, `menuAbout`）

### 步骤 2：获取脚本ID

1. 在 Google Apps Script 编辑器中，点击 **"文件"** → **"项目属性"**
2. 在 **"项目属性"** 对话框中，找到 **"脚本 ID"** 字段
3. 复制该脚本ID，这是库的唯一标识符

### 步骤 3：创建版本（可选）

如果需要使用固定版本（而不是开发模式），可以创建版本：

1. 在 Google Apps Script 编辑器中，点击 **"文件"** → **"管理版本"**
2. 在弹出的对话框中，输入版本说明（例如，"初始版本"）
3. 点击 **"保存新版本"**
4. 版本号会自动生成（1, 2, 3...）

**注意**：如果不创建版本，可以使用 **"开发模式"**，这样会自动使用最新代码。

### 步骤 4：发布为库（可选）

实际上，Google Apps Script 库不需要显式"发布"，只需要：
- 确保代码已保存
- 获取脚本ID
- 在其他项目中通过脚本ID添加库即可

如果需要更正式的方式，可以：
1. 点击 **"发布"** → **"部署为库"**（如果此选项存在）
2. 但通常不需要此步骤，直接使用脚本ID即可

## 在新表格中使用

### 步骤 1：添加库

1. 在 Google Sheets 中打开新表格
2. 点击 **"扩展程序"** → **"Apps Script"**
3. 点击 **"资源"** → **"库"**（或 **"库"** → **"添加库"**）
4. 在弹出的对话框中，粘贴之前复制的库的脚本ID
5. 点击 **"查找"** 或 **"添加"**
6. 在 **"标识符"** 字段中，输入一个简短的名称（例如：`CalendarSyncLib`），用于在代码中引用该库
7. 选择版本：
   - **"开发模式"**：自动使用最新代码（推荐用于开发）
   - **"版本号"**：使用固定版本（推荐用于生产环境）
8. 点击 **"添加"** 或 **"保存"**

### 步骤 2：创建初始化代码

在新表格的 Apps Script 编辑器中，创建以下代码：

```javascript
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
```

**重要说明**：
- 菜单项函数（`menuRunSync`, `menuViewConfig`, `menuViewStatus`, `menuAbout`）必须在用户表格中定义
- 这些函数作为包装函数，调用库中的实际函数
- 菜单系统无法直接找到库中的函数，所以需要这些包装函数

### 步骤 3：授权

1. 首次使用时，系统会提示授权
2. 点击 **"授权"** 并完成授权流程
3. 授予必要的权限：
   - Google Sheets 访问权限
   - Google Calendar 访问权限
   - 发送邮件权限（用于取消课程通知）

### 步骤 4：使用

1. 刷新表格页面
2. 应该能看到 **"📅 课程同步"** 菜单
3. 点击菜单中的 **"🔄 执行同步"** 开始处理课程数据

## 可用的库函数

### 主要函数

- **`createMenu()`** - 创建自定义菜单
  ```javascript
  CalendarSyncLib.createMenu();
  ```

- **`main()`** - 执行同步（主函数）
  ```javascript
  CalendarSyncLib.main();
  ```

### 菜单项函数

这些函数通常通过菜单调用，但也可以直接调用：

- **`menuRunSync()`** - 执行同步（带确认对话框）
- **`menuViewConfig()`** - 查看配置表
- **`menuViewStatus()`** - 查看状态表
- **`menuAbout()`** - 显示关于信息

## 配置要求

### 配置表（_SheetConfig）

在新表格中需要创建配置表 `_SheetConfig`，包含以下列：

| 列名 | 说明 | 必填 |
|------|------|------|
| Sheet名称 | 要处理的课程表名称 | 是 |
| 组织者日历ID | 组织者日历的ID | 是 |
| 老师邮箱 | 老师邮箱地址 | 否 |
| 学生邮箱 | 学生邮箱地址 | 否 |
| 时区 | 时区设置（如 Asia/Shanghai） | 否 |
| 提醒时间 | 提前提醒的分钟数 | 否 |

### 课程表

每个课程表需要包含以下列：

| 列名 | 说明 | 必填 |
|------|------|------|
| 课次 | 课程次数 | 是 |
| 日期 | 课程日期 | 是 |
| 开始时间 | 课程开始时间 | 是 |
| 结束时间 | 课程结束时间 | 是 |
| 课程内容 | 课程标题 | 是 |
| 老师姓名 | 老师姓名 | 否 |
| 学生姓名 | 学生姓名 | 否 |

## 注意事项

1. **权限**：首次使用时需要授权，确保授予所有必要的权限
2. **配置表**：必须创建配置表 `_SheetConfig`，否则无法运行
3. **日历ID**：确保组织者日历ID正确，并且有访问权限
4. **版本更新**：如果库更新了，可以选择更新到新版本
5. **开发模式**：建议使用开发模式，这样可以自动使用最新版本

## 故障排除

### 菜单不显示

- 检查是否创建了 `onOpen()` 函数
- 检查是否调用了 `CalendarSyncLib.createMenu()`
- 检查库是否正确添加
- 刷新表格页面

### 执行失败

- 检查配置表是否存在
- 检查配置表中的数据是否正确
- 检查日历ID是否正确
- 查看执行日志（扩展程序 → Apps Script → 执行日志）

### 权限问题

- 确保授予了所有必要的权限
- 检查日历是否有访问权限
- 检查表格是否有访问权限

## 更新库

如果库更新了：

1. 在 Apps Script 编辑器中，点击 **"资源"** → **"库"**
2. 找到已添加的库
3. 点击版本号旁边的下拉菜单
4. 选择新版本或 **"开发模式"**
5. 点击 **"保存"**

## 技术支持

如有问题，请查看：
- 用户手册：`USER_MANUAL.md`
- 技术设计文档：`TECHNICAL_DESIGN.md`
- README：`README.md`

