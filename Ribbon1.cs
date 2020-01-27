using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace vsto1
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {

            System.Windows.Forms.MessageBox.Show("it works ");
            string[] lines = { "Carmen i love you ", "Elisa i Love you", "Evelyn i Love you " };
            System.IO.File.WriteAllLines(@"C:\Users\mldvn\test\WriteLines.txt", lines);
        }
    }
}
