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
//using Microsoft.Toolkit.Uwp.Notifications;
//using Windows.UI.Notifications;
//using SaveAllTheHomework.Source.DesktopNotifications;
//using Windows.Data.Xml.Dom;

namespace SaveAllTheHomework.Source
{
    public class HomeworkBot
    {
        public HomeworkBot() {
            
        }

        public int SaveAllHomework() {



            Outlook.MAPIFolder ActivedFolder = Globals.ThisAddIn.Application.ActiveExplorer().CurrentFolder;

            List<HomeworkItem> HomeworkList = new List<HomeworkItem>();


            foreach (dynamic myItem in ActivedFolder.Items) {
                /// myItem.Sender.Address System.String
                /// myItem.Attachments.Count
                /// 
                /// myItem.CreationTime System.DateTime
                /// myItem.ConversationTopic System.String


                String sender = myItem.Sender.Address;

                HomeworkItem aHomeworkItem = new HomeworkItem
                {
                    Sender = sender
                };

                /// 识别学号
                // 尝试识别武汉大学邮箱凭此获得学号
                aHomeworkItem.StudentID = GetStudentIDFromWHUEmail(sender);

                // 尝试识别附件文件名凭此获得学号
                if (aHomeworkItem.StudentID == 0) {
                    foreach (dynamic myAttachment in myItem.Attachments)
                    {
                        aHomeworkItem.StudentID = GetStudentIDFromString(myAttachment.FileName);
                        if (aHomeworkItem.StudentID != 0) break;
                    }
                }

                // 尝试识别邮件标题凭此获得学号
                if (aHomeworkItem.StudentID == 0) {
                    aHomeworkItem.StudentID = GetStudentIDFromString(myItem.ConversationTopic);
                }

                // 邮件发送时间 
                aHomeworkItem.CreationTime = myItem.CreationTime;


                // 如果都识别不到
                if (aHomeworkItem.StudentID == 0) {
                    // 发出警告然后下一个
                    /*
                    // Construct the visuals of the toast (using Notifications library)
                    ToastContent toastContent = new ToastContent()
                    {
                        // Arguments when the user taps body of toast
                        Launch = "action=viewConversation&conversationId=5",

                        Visual = new ToastVisual()
                        {
                            BindingGeneric = new ToastBindingGeneric()
                            {
                                Children =
                                {
                                    new AdaptiveText()
                                    {
                                        Text = "Hello world!"
                                    }
                                }
                            }
                        }
                    };

                    // Create the XML document (BE SURE TO REFERENCE WINDOWS.DATA.XML.DOM)
                    var doc = new XmlDocument();
                    doc.LoadXml(toastContent.GetContent());

                    // And create the toast notification
                    var toast = new ToastNotification(doc);

                    // And then show it
                    DesktopNotificationManagerCompat.CreateToastNotifier().Show(toast);
                    */
                    MessageBox.Show("" + aHomeworkItem.StudentID);
                    continue;
                };


                if (aHomeworkItem.StudentID > 0)
                {
                    MessageBox.Show("" + aHomeworkItem.StudentID);
                }
                else {
                    MessageBox.Show("" + aHomeworkItem.StudentID);
                }

                HomeworkList.Add(aHomeworkItem);

                /*
                if (myItem.Attachments.Count > 1)
                {

                    foreach (dynamic myAttachment in myItem.Attachments)
                    {
                        myAttachment.SaveAsFile("D:\\AppData\\outlook\\" + myAttachment.FileName);
                    }

                }
                else
                {

                    foreach (dynamic myAttachment in myItem.Attachments)
                    {
                        myAttachment.SaveAsFile("D:\\AppData\\outlook\\" + myAttachment.FileName);
                    }

                }
                */

            }

            MessageBox.Show(ActivedFolder.FullFolderPath+" 里的所有附件都已经保存下来辣！");
           



            return 0;
        }

        private bool IsWhuEMail(String sender) {
            String[] Email = sender.Split('@');
            if (!Email[1].ToLower().Equals("whu.edu.cn")) return false;

            if (Email[0].Length != 13) return false; // 2018 302 114514 // 4+3+6

            for (int i = 0; i < 13; i++) {
                if ( (Email[0][i] < '0') || (Email[0][i] > '9') ) return false;
            }

            return true;
        }

        private long GetStudentIDFromWHUEmail(String sender) {
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
