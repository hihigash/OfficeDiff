namespace OfficeDiff
{
    public class WordComparer : IOfficeComparer
    {
        public void Compare(string orginal, string target)
        {
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application { Visible = true };
            word.DocumentOpen += doc =>
            {
                doc.Compare(target);
                doc.Close();
                word.Activate();

            };
            word.Documents.Open(orginal);
        }
    }
}