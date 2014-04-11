using System;

namespace Codesharper.PowerPoint.Helper.Objects
{
    using Microsoft.Office.Interop.PowerPoint;

    public class SlideComment
    {
        public string Author
        {
            get;
            set;
        }

        public string AuthorInitials
        {
            get;
            set;

        }

        public string Comment
        {
            get;
            set;
        }

        public float LeftPosition
        {
            get;
            set;
        }

        public float TopPosition
        {
            get;
            set;
        }
    }
}
