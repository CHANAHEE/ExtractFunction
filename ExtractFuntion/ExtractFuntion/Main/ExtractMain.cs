using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExtractFuntion
{
    public class ExtractMain
    {
        public bool Init()
        {
            if(ExcelManager.Instance.Init() == false)
            {
                return false;
            }

            MainForm NewMainForm = new MainForm();
            NewMainForm.Show();

            return true;
        }
    }
}
