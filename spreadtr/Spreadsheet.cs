namespace spreadstr
{
    using System;

    public class Spreadsheet
    {
        public Spreadsheet(byte[] content)
        {
            if (content == null)
                throw new ArgumentNullException("content");

            Content = content;
        }

        public byte[] Content { get; set; }
    }
}