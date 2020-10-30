using System;

namespace SampleApp.Model
{
    public class Book
    {
        public long Isbn { get; set; }

        public string Title { get; set; }

        public string Subtitle { get; set; }

        public string Author { get; set; }

        public DateTime Published { get; set; }

        public string Publisher { get; set; }

        public int Pages { get; set; }

        public string Description { get; set; }

        public string Website { get; set; }
    }
}
