import joplin from "api";
import { ContentScriptType, MenuItemLocation } from "api/types";
import { exec } from "child_process";
import * as fs from "fs";
import * as path from "path";
import * as os from "os";

joplin.plugins.register({
  onStart: async function () {
    await joplin.commands.register({
      name: "exportNoteToDocx",
      label: "导出笔记为Word文档",
      execute: async (noteId: string) => {
        try {
          // 获取当前笔记
          const note = await joplin.data.get(["notes", noteId], {
            fields: ["title", "body"],
          });

          // 创建临时目录用于存放转换过程中的文件
          const tempDir = os.tmpdir();
          const sanitizedTitle = note.title.replace(/[\\/:*?"<>|]/g, "_"); // 移除Windows文件名不允许的字符
          const markdownPath = path.join(tempDir, `${sanitizedTitle}.md`);

          // 显示选项对话框
          const dialog = await joplin.views.dialogs.create(
            "exportOptionsDialog"
          );
          await joplin.views.dialogs.setHtml(
            dialog,
            `  
            <form name="exportOptions">  
              <div style="padding: 10px;">  
                <h3>导出笔记为Word文档</h3>  
                <p>笔记标题: ${note.title}</p>  
                
                <div style="margin-bottom: 10px;">  
                  <label>  
                    <input type="checkbox" name="includeImages" checked>   
                    尝试包含图片  
                  </label>  
                </div>  
                
                <div style="margin-bottom: 10px;">  
                  <label>  
                    输出文件夹:  
                    <input type="text" name="outputFolder" value="${os
                      .homedir()
                      .replace(/\\/g, "\\\\")}" style="width: 100%;">   
                  </label>  
                  <small>留空则保存到桌面</small>  
                </div>  
              </div>  
            </form>  
          `
          );

          // 设置对话框按钮
          await joplin.views.dialogs.setButtons(dialog, [
            { id: "cancel", title: "取消" },
            { id: "ok", title: "导出" },
          ]);

          // 显示对话框并获取结果
          const result = await joplin.views.dialogs.open(dialog);

          if (result.id !== "ok") {
            return; // 用户取消
          }

          // 获取用户选择
          const includeImages = result.formData.exportOptions.includeImages;
          let outputFolder = result.formData.exportOptions.outputFolder.trim();

          // 如果输出文件夹为空，使用桌面
          if (!outputFolder) {
            outputFolder = path.join(os.homedir(), "Desktop");
          }

          // 确保输出文件夹存在
          if (!fs.existsSync(outputFolder)) {
            try {
              fs.mkdirSync(outputFolder, { recursive: true });
            } catch (error) {
              console.error(`创建输出文件夹失败:`, error);
              await joplin.views.dialogs.showMessageBox(
                `创建输出文件夹失败: ${
                  error instanceof Error ? error.message : String(error)
                }`
              );
              return;
            }
          }

          // 输出docx文件路径
          const docxPath = path.join(outputFolder, `${sanitizedTitle}.docx`);

          try {
            // 显示进度对话框
            await joplin.views.dialogs.showMessageBox("正在处理中，请稍候...");

            // 将笔记内容写入临时markdown文件
            fs.writeFileSync(markdownPath, note.body);

            // 构建pandoc命令
            let pandocCmd = "";

            if (includeImages) {
              // 如果包含图片，使用--extract-media参数
              const mediaDir = path.join(tempDir, `${sanitizedTitle}_media`);
              pandocCmd = `pandoc "${markdownPath}" -f markdown -t docx --extract-media="${mediaDir}" -o "${docxPath}"`;
            } else {
              pandocCmd = `pandoc "${markdownPath}" -f markdown -t docx -o "${docxPath}"`;
            }

            // 执行pandoc转换命令
            exec(pandocCmd, async (execError, stdout, stderr) => {
              // 无论成功与否，都删除临时markdown文件
              try {
                if (fs.existsSync(markdownPath)) {
                  fs.unlinkSync(markdownPath);
                }
              } catch (e) {
                console.error("删除临时文件失败:", e);
              }

              if (execError) {
                console.error(`执行错误:`, execError);
                await joplin.views.dialogs.showMessageBox(
                  `转换失败: ${execError.message}`
                );
                return;
              }

              if (stderr) {
                console.error(`标准错误: ${stderr}`);
              }

              // 检查docx文件是否成功创建
              if (fs.existsSync(docxPath)) {
                // 使用对话框API创建带按钮的对话框
                const confirmDialog = await joplin.views.dialogs.create(
                  "openFileDialog"
                );
                await joplin.views.dialogs.setHtml(
                  confirmDialog,
                  `  
                  <div style="padding: 10px;">  
                    <p>Word文档已成功导出到:</p>  
                    <p style="word-break: break-all; font-weight: bold;">${docxPath}</p>  
                    <p>是否立即打开?</p>  
                  </div>  
                `
                );

                await joplin.views.dialogs.setButtons(confirmDialog, [
                  { id: "open", title: "打开" },
                  { id: "close", title: "关闭" },
                ]);

                const dialogResult = await joplin.views.dialogs.open(
                  confirmDialog
                );

                // 如果用户选择打开文件
                if (dialogResult.id === "open") {
                  // 使用默认应用打开文件
                  const openCommand =
                    process.platform === "win32"
                      ? `start "" "${docxPath}"`
                      : process.platform === "darwin"
                      ? `open "${docxPath}"`
                      : `xdg-open "${docxPath}"`;

                  exec(openCommand, (openError) => {
                    if (openError) {
                      console.error("打开文件失败:", openError);
                      joplin.views.dialogs.showMessageBox(
                        `文件已保存，但无法自动打开: ${openError.message}`
                      );
                    }
                  });
                }
              } else {
                await joplin.views.dialogs.showMessageBox(
                  `转换似乎成功，但找不到输出文件: ${docxPath}`
                );
              }
            });
          } catch (error) {
            // 确保出错时也删除临时文件
            try {
              if (fs.existsSync(markdownPath)) {
                fs.unlinkSync(markdownPath);
              }
            } catch (e) {
              console.error("删除临时文件失败:", e);
            }

            console.error("执行pandoc命令时出错:", error);
            await joplin.views.dialogs.showMessageBox(
              `导出失败: ${
                error instanceof Error ? error.message : String(error)
              }`
            );
          }
        } catch (error) {
          console.error("执行插件过程中发生错误:", error);
          await joplin.views.dialogs.showMessageBox(
            `操作失败: ${
              error instanceof Error ? error.message : String(error)
            }`
          );
        }
      },
    });

    // 添加到笔记的上下文菜单
    await joplin.views.menuItems.create(
      "exportNoteToDocxMenuItem",
      "exportNoteToDocx",
      MenuItemLocation.NoteListContextMenu
    );

    // 也可以添加到笔记编辑器的上下文菜单
    await joplin.views.menuItems.create(
      "exportNoteToDocxEditorMenuItem",
      "exportNoteToDocx",
      MenuItemLocation.EditorContextMenu
    );

    // 添加到工具菜单
    await joplin.views.menuItems.create(
      "exportNoteToDocxToolsMenuItem",
      "exportNoteToDocx",
      MenuItemLocation.Tools
    );
  },
});
