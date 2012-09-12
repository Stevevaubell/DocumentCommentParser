using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentCommentParser.Logic.DTO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;

namespace DocumentCommentParser.Logic
{
    public class DocumentParser
    {
        public IList<DocumentComment> GetCommentsFromDocument(byte[] documentByteArray)
        {
            List<DocumentComment> docComments = new List<DocumentComment>();
            using (MemoryStream documentMemoryStream = new MemoryStream(documentByteArray))
            {
                using (WordprocessingDocument document = WordprocessingDocument.Open(documentMemoryStream, false))
                {
                    XDocument xmlDocument = document.MainDocumentPart.GetXDocument();

                    IList<XElement> documentElements = FilterOutXmlNodes(xmlDocument.Root);

                    string commentId = string.Empty;
                    foreach (Comment comment in document.MainDocumentPart.WordprocessingCommentsPart.Comments)
                    {
                        DocumentComment documentComment = new DocumentComment();

                        // Populate the other values from the comment object
                        documentComment.Author = comment.Author;
                        documentComment.Initials = comment.Initials;
                        documentComment.Date = comment.Date.HasValue ? comment.Date.Value : DateTime.Now;

                        // Populate the actual comment text into the object
                        foreach (Text text in comment.Descendants<Text>())
                        {
                            documentComment.CommentText += text.InnerText;
                        }

                        documentComment.CommentedText = GetCommentedOnText(documentElements, comment);

                        docComments.Add(documentComment);
                    }
                }
            }

            return docComments;
        }

        private static string GetCommentedOnText(IList<XElement> documentElements, Comment comment)
        {
            bool takeElements = false;
            StringBuilder commentedTextBuilder = new StringBuilder();

            // Get the commented on text
            foreach (XElement element in documentElements)
            {
                if (element.Name == W.commentRangeStart && element.Attribute(W.id).Value == comment.Id.Value)
                {
                    takeElements = true;
                    continue;
                }

                if (element.Name == W.commentRangeEnd && element.Attribute(W.id).Value == comment.Id.Value)
                {
                    takeElements = false;
                    break;
                }

                bool mustBreak = false;
                foreach (XElement subElement in element.Descendants())
                {
                    if (subElement.Name == W.commentRangeStart && subElement.Attribute(W.id).Value == comment.Id.Value)
                    {
                        takeElements = true;
                        continue;
                    }

                    if (subElement.Name == W.commentRangeEnd && subElement.Attribute(W.id).Value == comment.Id.Value)
                    {
                        takeElements = false;
                        mustBreak = true;
                        break;
                    }

                    if (takeElements && subElement.Name == W.t)
                    {
                        commentedTextBuilder.Append(subElement.Value.ToString());
                    }
                }

                if (mustBreak)
                {
                    break;
                }

                if (takeElements && element.Name == W.p)
                {
                    commentedTextBuilder.AppendLine();
                }
            }

            return commentedTextBuilder.ToString();
        }

        private static IList<XElement> FilterOutXmlNodes(XElement xmlElement)
        {
            List<XElement> elementsFound = xmlElement.Descendants().Where(
                e => (e.Name == W.p && e.Descendants().Any(ed => ed.Name == W.t || ed.Name == W.commentRangeStart || ed.Name == W.commentRangeEnd))
                    || e.Name == W.commentRangeStart || e.Name == W.commentRangeEnd).ToList();

            return elementsFound;
        }
    }
}
