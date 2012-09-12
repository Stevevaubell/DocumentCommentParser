using System;

namespace DocumentCommentParser.Logic.DTO
{
    public class DocumentComment
    {
        public string CommentId { get; set; }

        public string Author { get; set; }

        public DateTime Date { get; set; }

        public string Initials { get; set; }

        public string CommentText { get; set; }

        public string CommentedText { get; set; }
    }
}
