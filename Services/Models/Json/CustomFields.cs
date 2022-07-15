using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Services.Models.Json
{
    public class CustomFields
    {
        public string Id { get; set; }

        public string name { get; set; }

        public string value { get; set; }


        public static CustomFields  SetCustomFields(string id, string name, string value)

        {
            return new CustomFields() { Id = id, name = name, value = value };

        }


    }
}
