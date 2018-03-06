using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace NPOI.HWPF
{
    public  class Class1
    {


        public void test()
        {

            StringBuilder sb = new StringBuilder();
            using (FileStream stream = File.OpenRead("d:/test.doc"))
            {
                HWPFDocument hd = new HWPFDocument(stream);
                var table = hd.ParagraphTable;
               
                


            }
        }

    }
}
