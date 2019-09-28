using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Collections;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.IO;

namespace SaveAllTheHomework.Source
{
    public class HomeworkBot
    {
        public HomeworkBot()
        {

        }

        public int SaveAllHomework()
        {



            Outlook.MAPIFolder ActivedFolder = Globals.ThisAddIn.Application.ActiveExplorer().CurrentFolder;

            List<HomeworkItem> HomeworkList = new List<HomeworkItem>();


            foreach (dynamic myItem in ActivedFolder.Items)
            {
                /// myItem.Sender.Address System.String
                /// myItem.Attachments.Count
                /// 
                /// myItem.CreationTime System.DateTime
                /// myItem.ConversationTopic System.String
                /// SentOn

                String sender = "";

                /// 尝试解析发件人
                try
                {
                    sender = myItem.SenderEmailAddress;
                }
                catch (Exception e1) {
                    try
                    {
                        sender = myItem.Sender.Address;
                    }
                    catch (Exception e2)
                    {

                    };
                };


                if (sender.Equals("")) {
                    continue;
                }

                HomeworkItem aHomeworkItem = new HomeworkItem
                {
                    Sender = sender
                };

                /// 识别学号
                // 尝试识别武汉大学邮箱凭此获得学号
                aHomeworkItem.StudentID = GetStudentIDFromWHUEmail(sender);

                // 尝试识别附件文件名凭此获得学号
                if (aHomeworkItem.StudentID == 0)
                {
                    foreach (dynamic myAttachment in myItem.Attachments)
                    {
                        aHomeworkItem.StudentID = GetStudentIDFromString(myAttachment.FileName);
                        if (aHomeworkItem.StudentID != 0) break;
                    }
                }

                // 尝试识别邮件标题凭此获得学号
                if (aHomeworkItem.StudentID == 0)
                {
                    aHomeworkItem.StudentID = GetStudentIDFromString(myItem.ConversationTopic);
                }

                // 如果都识别不到
                if (aHomeworkItem.StudentID == 0)
                {
                    continue;
                };

                /// 邮件发送时间 
                aHomeworkItem.SentOn = myItem.SentOn;

                // 如果没有附件
                if (myItem.Attachments.Count == 0)
                {
                    continue;
                }

                /// 获取附件
                aHomeworkItem.Attachments = myItem.Attachments;

                HomeworkList.Add(aHomeworkItem);


            }

            // 对所有邮件按照
            //  学号第一次序升序
            //  邮件发送时间第二次徐升序 排序
            HomeworkList.Sort();

            for (int i=0; i<HomeworkList.Count(); i++) {
                //foreach (HomeworkItem k in HomeworkList) {

                // 如果相同学号存在更新版本的作业则跳过旧版本作业
                if (i < HomeworkList.Count() - 1) {
                    if (HomeworkList[i].StudentID == HomeworkList[i + 1].StudentID) continue;
                }

                HomeworkItem k = HomeworkList[i];

                // 将附件保存
                String RootFolder = "D:\\AppData\\outlook\\";
                if (!Directory.Exists(RootFolder)) Directory.CreateDirectory(RootFolder);

                if (k.Attachments.Count == 1)
                {
                    foreach (dynamic myAttachment in k.Attachments)
                    {
                        String FileName = myAttachment.FileName;
                        String[] FileNames = FileName.Split('.');

                        String OutFileName = k.StudentID + "." + FileNames[FileNames.Count() - 1];

                        if (OutFileName.Equals(FileName))
                        {
                            myAttachment.SaveAsFile(RootFolder + FileName);
                        }
                        else
                        {
                            String TempFolder = RootFolder + k.StudentID + "\\";
                            if (!Directory.Exists(TempFolder)) Directory.CreateDirectory(TempFolder);
                            myAttachment.SaveAsFile(TempFolder + FileName);
                        }
                    }
                }
                else {
                    foreach (dynamic myAttachment in k.Attachments)
                    {
                        String TempFolder = RootFolder + k.StudentID + "\\";
                        if (!Directory.Exists(TempFolder)) Directory.CreateDirectory(TempFolder);
                        myAttachment.SaveAsFile(TempFolder + myAttachment.FileName);

                    }
                }

            }




            MessageBox.Show(ActivedFolder.FullFolderPath + " 里的所有附件都已经保存下来辣！");




            return 0;
        }

        private bool IsWhuEMail(String sender)
        {
            String[] Email = sender.Split('@');
            if (!Email[1].ToLower().Equals("whu.edu.cn")) return false;

            if (Email[0].Length != 13) return false; // 2018 302 114514 // 4+3+6

            for (int i = 0; i < 13; i++)
            {
                if ((Email[0][i] < '0') || (Email[0][i] > '9')) return false;
            }

            return true;
        }

        private long GetStudentIDFromWHUEmail(String sender)
        {
            String[] Email = sender.Split('@');
            if (!Email[1].ToLower().Equals("whu.edu.cn")) return 0;
            if (Email[0].Length != 13) return 0; // 2018 302 114514 // 4+3+6

            for (int i = 0; i < 13; i++)
            {
                if ((Email[0][i] < '0') || (Email[0][i] > '9')) return 0;
            }

            return long.Parse(Email[0]);
        }

        private long GetStudentIDFromString(String sender)
        {
            string pattern = @"\d{13}";

            foreach (Match match in Regex.Matches(sender, pattern))
                return long.Parse(match.Value);

            return 0;
        }
    }
}
