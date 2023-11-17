using Xceed.Document.NET;
using Xceed.Words.NET;

namespace KaitReferences.Extensions
{
    public static class DocXExtension
    {
        static Formatting format = new Formatting()
        {
            FontFamily = new Font("Times New Roman"),
            Size = 14
        };

        public static void SetText(this DocX document, string bookmark, string text, Formatting? formating = null)
        {
            formating ??= format;
            document.Bookmarks[bookmark].SetText(text, formating);
        }

        public static void WithUnderline(this DocX document, bool isTrue)
        {
            format.UnderlineStyle = isTrue ? UnderlineStyle.singleLine : UnderlineStyle.none;
        }
    }
}
