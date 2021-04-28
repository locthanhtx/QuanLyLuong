using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SendMail.Model
{
    class InfoUser
    {
        public string hoTen;
        public List<string> email;
        public List<string> fileName;
        public string content;
        public string subject;
        public InfoUser()
        {
            hoTen = string.Empty;
            content = string.Empty;
            subject = string.Empty;
            email = new List<string>();
            fileName = new List<string>();
        }
    }
}
