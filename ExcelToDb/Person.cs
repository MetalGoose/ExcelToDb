using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelToDb
{
    internal struct Person
    {
        public Guid Id { get; set; }

        public string PersonInitials { get; set; }

        public string Gender { get; set; }

        public string Age { get; set; }

        public string PersonCode { get; set; }

        public string Class { get; set; }

        public string SchoolNum { get; set; }

        public string City { get; set; }

        public string FirstTesterInitials { get; set; }
        public string SecondTesterInitials { get; set; }

        public string OnlineOrOffline { get; set; }

        public DateTime TestDate { get; set; }

        public override string ToString()
        {
            return string.Format(
                $"{Id}_{PersonInitials}_{Gender}_{Age}_{Class}_{SchoolNum}_{City}_{FirstTesterInitials}{SecondTesterInitials}_{PersonCode}_{OnlineOrOffline}_{TestDate}");
        }
    }
}