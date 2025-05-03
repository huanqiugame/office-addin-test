/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // 初始化Excel事件监听
  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    
    // 使用createProxy方式绑定事件处理
    sheet.onChanged.add(async (event) => {
      try {
        await Excel.run(async (innerContext) => {
          const range = innerContext.workbook.worksheets.getItem(event.worksheetId).getRange(event.address);
          range.load('values,format');
          await innerContext.sync();

          const text = range.values[0][0]?.toString() || '';
          // 优化正则表达式以匹配实时输入
          const headingMatch = /^\s*(#{1,6})(?=\s)/.exec(text);

          if (headingMatch) {
            const level = Math.min(6, headingMatch[1].length);
            const className = `heading-${level}`;
            
            // 强制清除所有现有格式
            range.format
              .borders.removeAll()
              .font.set({ bold: false, italic: false, underline: false, size: 11 });
            
            // 应用新标题类
            range.format.cssClass = className;
            
            // 保持光标位置不变
            innerContext.workbook.application.activeCell = range;
          }
          
          await innerContext.sync();
        });
      } catch (error) {
        console.error('Markdown处理错误:', error);
      }
    });
    
    await context.sync();
  });
});

// Register the function with Office.
Office.actions.associate("action", () => {}); // 保留空操作，因为我们使用事件监听