using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Collections;
using System.Windows.Forms;

namespace SaveAllTheHomework.Source
{
    public class HomeworkBot
    {
        public HomeworkBot() {
            
        }

        public int SaveAllHomework() {
            Outlook.MAPIFolder ActivedFolder = Globals.ThisAddIn.Application.ActiveExplorer().CurrentFolder;
            
            foreach (dynamic myItem in ActivedFolder.Items) {
                /// myItem.Sender.Address System.String
                /// myItem.Attachments.Count
                /// 

                foreach (dynamic myAttachment in myItem.Attachments) {
                    myAttachment.SaveAsFile("D:\\AppData\\outlook\\"+ myAttachment.FileName);
                }

            }

            MessageBox.Show(ActivedFolder.FullFolderPath+" 里的所有附件都已经保存下来辣！");
           



            return 0;
        }
    }
}
