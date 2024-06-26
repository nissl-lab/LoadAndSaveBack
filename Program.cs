﻿using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace LoadAndSaveBack
{
    class Program
    {
        enum RunMode
        { 
            Excel2003,
            Excel2007,
            Word
        }
        static void Main(string[] args)
        {
            if (args.Length != 2)
                return;

            string src = args[0];
            string target = args[1];

            RunMode mode= RunMode.Excel2007;
            if (src.EndsWith(".docx"))
                mode = RunMode.Word;
            else if(src.EndsWith(".xls"))
                mode = RunMode.Excel2003;

            if (mode == RunMode.Excel2007)
            {
                using (Stream rfs = File.OpenRead(src))
                {
                    using (IWorkbook workbook = new XSSFWorkbook(rfs))
                    {
                        using (FileStream fs = File.Create(target))
                        {
                            workbook.Write(fs, false);
                        }
                    }
                }
            }
            else if (mode== RunMode.Excel2003)
            {
                using (Stream rfs = File.OpenRead(src))
                {
                    using (IWorkbook workbook = new HSSFWorkbook(rfs))
                    {
                        using (FileStream fs = File.Create(target))
                        {
                            workbook.Write(fs, false);
                        }
                    }
                }
            }
            else
            {
                using (Stream rfs = File.OpenRead(src))
                {
                    using (XWPFDocument workbook = new XWPFDocument(rfs))
                    {
                        using (FileStream fs = File.Create(target))
                        {
                            workbook.Write(fs);
                        }
                    }
                }
            }
        }
    }
}
