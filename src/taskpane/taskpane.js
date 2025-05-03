/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// 将Office初始化移到文档加载完成后
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initializeOffice);
} else {
  initializeOffice();
}

// 直接在window对象上初始化属性
window.markdownTrigger = null;

function initializeOffice() {
  Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
      // 安全移除sideload消息
      const sidenavMsg = document.getElementById("sideload-msg");
      if (sidenavMsg && sidenavMsg.parentNode) {
        sidenavMsg.parentNode.removeChild(sidenavMsg);
      }
      
      const appBody = document.getElementById("app-body");
      if (appBody) appBody.style.display = "block";
      
      initMarkdownListener();
    }
  });
}

async function createTable() {
  await Excel.run(async (context) => {

      // TODO1: Queue table creation logic here.
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      const expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
      expensesTable.name = "ExpensesTable";
      // TODO2: Queue commands to populate the table with data.

      // TODO3: Queue commands to format the table.

      await context.sync();
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
      await callback();
  } catch (error) {
      // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
      console.error(error);
  }
}

// 将键盘监听移到DOM加载完成后
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initMarkdownListener);
} else {
  initMarkdownListener();
}

// 使用类型安全的定时器
function initMarkdownListener() {
  // 使用闭包变量替代window属性
  let markdownTrigger = null;
  
  document.addEventListener('keydown', (event) => {
    if (event.key === ' ' && !event.shiftKey && !event.ctrlKey) {
      // 清除之前的定时器
      if (markdownTrigger) {
        clearTimeout(markdownTrigger);
      }
      
      // 设置新的定时器
      markdownTrigger = setTimeout(() => {
        checkMarkdownFormat();
        markdownTrigger = null;
      }, 150);
    }
  });
}

async function checkMarkdownFormat() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getActiveCell();
      range.load('values,format');
      await context.sync();

      // 安全处理单元格值
      const rawValue = range.values[0][0];
      const text = rawValue !== null && rawValue !== undefined 
        ? rawValue.toString() 
        : '';

      const headingMatch = /^\s*(#{1,6})(?=\s)/.exec(text);

      if (headingMatch) {
        const level = Math.min(6, headingMatch[1].length);
        const className = `heading-${level}`;
        
        // 使用格式化方法替代cssClass
        applyHeadingStyle(range, level);
        
        await context.sync();
      }
    });
  } catch (error) {
    console.error('样式应用失败:', error);
  }
}

// 简化样式应用逻辑
function applyHeadingStyle(range, level) {
  const styles = {
    1: { fontSize: 24, bold: true, color: '#2C3E50' },
    2: { fontSize: 20, bold: true, color: '#34495E' },
    3: { fontSize: 18, bold: true, color: '#34495E' },
    4: { fontSize: 16, bold: true, color: '#2C3E50' },
    5: { fontSize: 14, bold: true, color: '#34495E' },
    6: { fontSize: 12, bold: true, color: '#7F8C8D' }
  };
  
  const style = styles[level] || styles[1];
  
  // 安全应用样式
  try {
    range.format
      .borders.removeAll()
      .font.set({
        bold: style.bold,
        size: style.fontSize,
        color: style.color
      });
  } catch (error) {
    console.error('样式应用失败:', error);
  }
}
