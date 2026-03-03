using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace jimsoutlooktools
{
    public partial class RibbonTools
    {
        private const string AppVersion = "v1.0.3";

        private void RibbonTools_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void btnSaveAttachments_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                string saveRoot;
                DateTime startDate, endDate;

                if (!SelectSaveOptions(out saveRoot, out startDate, out endDate))
                {
                    MessageBox.Show("操作取消。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                Outlook.MAPIFolder inbox = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                Outlook.Items items = inbox.Items;
                // 限制只获取必要字段，减少内存占用
                items.IncludeRecurrences = false;

                int savedCount = 0;
                int skippedCount = 0;
                int processedCount = 0;
                var failedAttachments = new List<string>();

                using (var progressForm = new ProgressForm(AppVersion))
                {
                    progressForm.Show();
                    progressForm.SetProgress(0, items.Count);

                    // 使用 for 循环代替 foreach，更好地控制 COM 对象释放
                    for (int i = 1; i <= items.Count; i++)
                    {
                        object item = null;
                        Outlook.MailItem mailItem = null;
                        Outlook.Attachments attachments = null;

                        try
                        {
                            item = items[i];
                            mailItem = item as Outlook.MailItem;

                            if (mailItem != null && mailItem.ReceivedTime >= startDate && mailItem.ReceivedTime <= endDate)
                            {
                                string monthFolder = Path.Combine(saveRoot, mailItem.ReceivedTime.ToString("yyyyMM"));
                                Directory.CreateDirectory(monthFolder);

                                attachments = mailItem.Attachments;
                                for (int j = 1; j <= attachments.Count; j++)
                                {
                                    Outlook.Attachment attachment = null;
                                    try
                                    {
                                        attachment = attachments[j];

                                        // 跳过内联图片（小于100KB的图片文件通常是邮件正文中的图标、表情等）
                                        string ext = Path.GetExtension(attachment.FileName).ToLower();
                                        bool isImage = ext == ".png" || ext == ".jpg" || ext == ".jpeg" ||
                                                       ext == ".gif" || ext == ".bmp" || ext == ".ico" || ext == ".webp";

                                        if (isImage && attachment.Size < 102400) // 小于100KB的图片跳过
                                        {
                                            continue;
                                        }

                                        string safeFileName = SanitizeFileName(attachment.FileName);
                                        // 使用邮件接收时间戳+原文件名作为唯一标识
                                        string timestamp = mailItem.ReceivedTime.ToString("yyyyMMdd_HHmmss_fff");
                                        string uniqueFileName = $"{timestamp}_{safeFileName}";
                                        string targetPath = Path.Combine(monthFolder, uniqueFileName);

                                        if (File.Exists(targetPath))
                                        {
                                            skippedCount++;
                                        }
                                        else
                                        {
                                            // 单独捕获每个附件保存的异常
                                            try
                                            {
                                                attachment.SaveAsFile(targetPath);
                                                savedCount++;
                                            }
                                            catch (Exception ex)
                                            {
                                                // 记录失败的附件信息
                                                string failedInfo = $"文件: {attachment.FileName} | 邮件: {mailItem.Subject} | 时间: {mailItem.ReceivedTime:yyyy-MM-dd HH:mm:ss} | 错误: {ex.Message}";
                                                failedAttachments.Add(failedInfo);
                                                System.Diagnostics.Debug.WriteLine($"保存附件失败: {failedInfo}");
                                            }
                                        }
                                    }
                                    finally
                                    {
                                        // 释放附件 COM 对象
                                        if (attachment != null)
                                        {
                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(attachment);
                                        }
                                    }
                                }
                                processedCount++;

                                // 每处理 50 封邮件强制垃圾回收一次
                                if (processedCount % 50 == 0)
                                {
                                    GC.Collect();
                                    GC.WaitForPendingFinalizers();
                                }
                            }
                        }
                        finally
                        {
                            // 释放 COM 对象
                            if (attachments != null)
                            {
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(attachments);
                            }
                            if (mailItem != null)
                            {
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(mailItem);
                            }
                            if (item != null)
                            {
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(item);
                            }
                        }

                        progressForm.SetProgress(i, items.Count);

                        // 每 100 封邮件让 UI 刷新一下，避免假死
                        if (i % 100 == 0)
                        {
                            System.Windows.Forms.Application.DoEvents();
                        }
                    }
                }

                // 显示详细的保存结果
                ShowSaveResult(savedCount, skippedCount, failedAttachments);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"发生错误: {ex.Message}", $"jimsoutlooktools {AppVersion}", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShowSaveResult(int savedCount, int skippedCount, List<string> failedAttachments)
        {
            int failedCount = failedAttachments.Count;
            StringBuilder message = new StringBuilder();
            message.AppendLine($"保存完成！");
            message.AppendLine();
            message.AppendLine($"✓ 已保存: {savedCount} 个附件");
            message.AppendLine($"○ 跳过(已存在): {skippedCount} 个附件");
            message.AppendLine($"✗ 保存失败: {failedCount} 个附件");

            if (failedCount > 0)
            {
                message.AppendLine();
                message.AppendLine("失败详情:");
                message.AppendLine("--------------------");
                foreach (var failed in failedAttachments)
                {
                    message.AppendLine($"• {failed}");
                }
            }

            // 如果失败数量较多，使用滚动文本框显示
            if (failedCount > 5)
            {
                using (var resultForm = new SaveResultForm(AppVersion, message.ToString()))
                {
                    resultForm.ShowDialog();
                }
            }
            else
            {
                MessageBox.Show(message.ToString(), $"jimsoutlooktools {AppVersion} - 保存结果", 
                    MessageBoxButtons.OK, failedCount > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
            }
        }

        private bool SelectSaveOptions(out string saveRoot, out DateTime startDate, out DateTime endDate)
        {
            saveRoot = null;
            startDate = DateTime.MinValue;
            endDate = DateTime.MaxValue;

            using (var form = new DateRangePickerForm(AppVersion))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    saveRoot = form.SavePath;
                    // 起始日期设为当天00:00:00，结束日期设为当天23:59:59，确保包含整天
                    startDate = form.StartDate.Date;
                    endDate = form.EndDate.Date.AddHours(23).AddMinutes(59).AddSeconds(59);

                    if (startDate > endDate)
                    {
                        MessageBox.Show("起始日期不能晚于结束日期！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return false;
                    }

                    return true;
                }
            }

            return false;
        }

        private string SanitizeFileName(string fileName)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                fileName = fileName.Replace(c, '-');
            }

            return fileName.Length > 180 ? fileName.Substring(0, 180) : fileName;
        }
    }
}
