using System;
using System.Collections.Generic;

namespace SaveAllTheHomework.Source
{
    internal class HomeworkItem : IComparable, IComparer<HomeworkItem>
    {
        public String Sender = "";
        public long StudentID = 0;
        public dynamic Attachments = null;
        public DateTime SentOn = DateTime.UtcNow;

        public int Compare(HomeworkItem x, HomeworkItem y)
        {
            if (x.StudentID < y.StudentID) return -1;
            if (x.StudentID > y.StudentID) return 1;
            if (x.SentOn > y.SentOn) return 1;
            if (x.SentOn < y.SentOn) return -1;

            return 0;
        }

        public int CompareTo(object obj)
        {
            return Compare(this, (HomeworkItem)obj);
        }

    }


}