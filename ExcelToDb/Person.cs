using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelToDb
{
    internal class Person
    {
        public int Id { get; set; }

        public string PersonInitials { get; set; }

        public string Gender { get; set; }

        public string Age { get; set; }

        public string Class { get; set; }

        public string SchoolNum { get; set; }

        public string City { get; set; }

        public override string ToString()
        {
            return string.Format(
                $"{Id}_{PersonInitials}_{Gender}_{Age}_{Class}_{SchoolNum}_{City}");
        }
    }
}