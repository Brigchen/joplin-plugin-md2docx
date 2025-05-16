import joplin from "api";
import {
  ContentScriptType,
  MenuItemLocation,
  SettingItemType,
} from "api/types";
import { exec } from "child_process";
import * as fs from "fs";
import * as path from "path";
import * as os from "os";

// 添加对话框计数器以确保唯一 ID
let dialogCounter = 0;

joplin.plugins.register({
  onStart: async function () {
    // 创建专用的设置部分
    await joplin.settings.registerSection("md2docxSettings", {
      label: "md2docx-pandoc插件设置",
      iconName: "fas fa-file-word",
      description: "配置Markdown导出为Word文档的选项",
    });

    // 注册插件设置 - 使用正确的方法名 registerSettings
    await joplin.settings.registerSettings({
      defaultTemplatePath: {
        value: "",
        type: SettingItemType.String,
        section: "md2docxSettings",
        public: true,
        label: "Word模板文件路径",
        description:
          "请输入Word模板文件的完整路径，例如: C:\\Templates\\template.docx",
      },
      includeImagesDefault: {
        value: true,
        type: SettingItemType.Bool,
        section: "md2docxSettings",
        public: true,
        label: "默认包含图片",
        description: "勾选后，导出的Word文档将默认包含图片",
      },
      defaultOutputFolder: {
        value: path.join(os.homedir(), "Desktop"),
        type: SettingItemType.String,
        section: "md2docxSettings",
        public: true,
        label: "默认输出文件夹",
        description: "请输入默认输出文件夹的完整路径",
      },
    });

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
          const markdownPath = path.join(
            tempDir,
            `${sanitizedTitle}-${Date.now()}.md`
          );

          // 从设置中获取默认值
          const defaultTemplatePath = await joplin.settings.value(
            "defaultTemplatePath"
          );
          const includeImagesDefault = await joplin.settings.value(
            "includeImagesDefault"
          );
          const defaultOutputFolder = await joplin.settings.value(
            "defaultOutputFolder"
          );

          // 创建对话框，使用唯一ID避免冲突
          const dialogId = `exportOptionsDialog-${Date.now()}-${dialogCounter++}`;
          const dialog = await joplin.views.dialogs.create(dialogId);

          // 预定义一些常用文件夹
          const commonFolders = [
            os.homedir(),
            path.join(os.homedir(), "Desktop"),
            path.join(os.homedir(), "Documents"),
            // 可以根据需要添加更多常用路径
          ];

          // 设置HTML内容，使用默认设置值
          await joplin.views.dialogs.setHtml(
            dialog,
            `  
            <form name="exportOptions">  
              <div style="padding: 10px;">  
                <h3>导出笔记为Word文档</h3>  
                <p>笔记标题: ${note.title}</p>  
                
                <div style="margin-bottom: 10px;">  
                  <label>  
                    <input type="checkbox" name="includeImages" ${
                      includeImagesDefault ? "checked" : ""
                    }>   
                    尝试包含图片  
                  </label>  
                </div>  
                
                <div style="margin-bottom: 10px;">  
                  <label>  
                    <input type="checkbox" name="useTemplate" ${
                      defaultTemplatePath ? "checked" : ""
                    }>   
                    使用Word文档模板  
                  </label>  
                </div>  
                
                <div style="margin-bottom: 10px; margin-left: 20px;">  
                  <label>  
                    模板文件路径 (Word .docx文件):  
                    <input type="text" name="templatePath" value="${
                      defaultTemplatePath || ""
                    }" style="width: 100%;">   
                  </label>  
                  <small>  
                    使用插件设置中的模板文件，或输入其他模板的完整路径<br>  
                    <a href="#" onclick="window.open(':/plugins/settings?settingSection=md2docxSettings')">点击这里打开设置</a>  
                  </small>  
                </div>  
                
                <div style="margin-bottom: 10px;">  
                  <label>  
                    输出文件夹:  
                    <input type="text" id="outputFolder" name="outputFolder" value="${
                      defaultOutputFolder || ""
                    }" style="width: 100%;">   
                  </label>  
                  <small>留空则保存到桌面</small>  
                </div>  
                
                <div style="margin-bottom: 10px;">  
                  <label>  
                    常用文件夹:  
                    <select onchange="document.getElementById('outputFolder').value = this.value">  
                      <option value="">-- 选择常用文件夹 --</option>  
                      ${commonFolders
                        .map(
                          (folder) =>
                            `<option value="${folder.replace(
                              /\\/g,
                              "\\\\"
                            )}">${folder}</option>`
                        )
                        .join("")}  
                    </select>  
                  </label>  
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
          const useTemplate = result.formData.exportOptions.useTemplate;
          const templatePath = useTemplate
            ? result.formData.exportOptions.templatePath.trim()
            : "";
          let outputFolder = result.formData.exportOptions.outputFolder.trim();

          // 如果输出文件夹为空，使用桌面
          if (!outputFolder) {
            outputFolder = path.join(os.homedir(), "Desktop");
          }

          // 保存用户选择作为新的默认值
          await joplin.settings.setValue("defaultTemplatePath", templatePath);
          await joplin.settings.setValue("includeImagesDefault", includeImages);
          await joplin.settings.setValue("defaultOutputFolder", outputFolder);

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
            // 将笔记内容写入临时markdown文件
            fs.writeFileSync(markdownPath, note.body);

            // 构建pandoc命令
            let pandocCmd = "";

            if (useTemplate && templatePath && fs.existsSync(templatePath)) {
              // 如果使用模板，添加--reference-doc参数
              if (includeImages) {
                // 包含图片 + 使用模板
                const mediaDir = path.join(tempDir, `${sanitizedTitle}_media`);
                pandocCmd = `pandoc "${markdownPath}" -f markdown -t docx --reference-doc="${templatePath}" --extract-media="${mediaDir}" -o "${docxPath}"`;
              } else {
                // 仅使用模板
                pandocCmd = `pandoc "${markdownPath}" -f markdown -t docx --reference-doc="${templatePath}" -o "${docxPath}"`;
              }
            } else {
              // 不使用模板的情况
              if (includeImages) {
                // 仅包含图片
                const mediaDir = path.join(tempDir, `${sanitizedTitle}_media`);
                pandocCmd = `pandoc "${markdownPath}" -f markdown -t docx --extract-media="${mediaDir}" -o "${docxPath}"`;
              } else {
                // 基本转换
                pandocCmd = `pandoc "${markdownPath}" -f markdown -t docx -o "${docxPath}"`;
              }
            }

            console.log("执行命令:", pandocCmd);

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
                // 创建确认对话框，使用唯一ID
                const confirmId = `openFileDialog-${Date.now()}-${dialogCounter++}`;
                const confirmDialog = await joplin.views.dialogs.create(
                  confirmId
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

    // 注册命令以打开设置
    await joplin.commands.register({
      name: "openMd2DocxSettings",
      label: "md2docx-pandoc插件设置",
      execute: async () => {
        await joplin.commands.execute("openSettings", "md2docxSettings");
      },
    });

    // // 添加设置命令到工具菜单
    // await joplin.views.menuItems.create(
    //   "openMd2DocxSettingsMenuItem",
    //   "openMd2DocxSettings",
    //   MenuItemLocation.Tools
    // );
  },
});
