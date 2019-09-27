using System;

namespace SaveAllTheHomework.Source
{
    internal class HomeworkItem
    {
        public String Sender = "";
        public long StudentID = 0;
        public dynamic Attachments = null;
        public DateTime CreationTime = DateTime.UtcNow;
    }
}