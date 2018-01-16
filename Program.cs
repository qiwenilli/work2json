using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using System.Collections;
using Newtonsoft.Json;
using static System.Net.WebRequestMethods;
using System.IO;

// '/d/Program Files (x86)/Microsoft/ILMerge/ILMerge.exe' /target:winexe /targetplatform:v4 /out:a.exe ConsoleAppCredit.exe  Newtonsoft.Json.dll
namespace ConsoleAppCredit
{
    class Program
    {
        static void Main(string[] args)
        {
            // 0 保存的目录
            // 1 传入pdf文件名
            // 2 保存为的json文件名

            Console.WriteLine("传入所有参数");
            Console.WriteLine(JsonConvert.SerializeObject(args));

            //
            string fsource = args[0];
            string fdes = args[1];
            
            if (args.Length !=2)
            {
                Console.WriteLine("args length not is 2");

                return;
            }

            docProcess(fsource, fdes);
           
        }

        static void docProcess(string filename, string savejsontxtfile)
        {
            Word.Application wordApp = new Word.Application();
            Word.Document wordDoc = new Word.Document();

            if (!System.IO.File.Exists(filename))
            {
                Console.WriteLine(filename + " 文件不存在");
                return;
            }

            try
            {
                wordDoc = wordApp.Documents.Open(filename);
            }
            catch(Exception err)
            {
                Console.WriteLine("pdf 文件打开异常");
                return;
            }
            wordApp.Visible = false;


            List<String[,]> tableDataList = getTableDataList(wordDoc);
            List<Word.Range> TablesRanges = getTableRangeList(wordDoc);


            //遍历所有段落，找出表格
            ArrayList docDataList = new ArrayList();
            //
            Boolean bInTable;
            //
            int table_i = -1;
            int doc_i = 0;
            String[,] part_data = new String[1, 1];
            //
            string part_string;
            for (int pi = 0; pi <= wordDoc.Paragraphs.Count; pi++)
            {
                bInTable = false;
                try
                {
                    Word.Range r = wordDoc.Paragraphs[pi].Range;

                    part_string = r.Text;

                    //Console.WriteLine(">>");
                    //Console.WriteLine(part_string);

                    for (int i = 0; i < TablesRanges.Count; i++)
                    {
                        if (r.Start >= TablesRanges[i].Start && r.Start < TablesRanges[i].End)
                        {
                            //Console.WriteLine(">>>");
                            //Console.WriteLine(part_string);

                            bInTable = true;
                            if (i > table_i)
                            {
                                table_i = i;
                            }
                            break;
                        }
                    }

                    if (bInTable == false)
                    {
                        //Console.WriteLine("part---::::");
                        //Console.WriteLine(part_string);

                        docDataList.Add(part_string);
                    }
                    else
                    {
                        if (table_i > doc_i)
                        {
                            docDataList.Add(tableDataList[table_i]);

                            doc_i = table_i;
                        }
                        bInTable = false;
                    }

                }
                catch (Exception ee)
                {
                    Console.WriteLine("解析失败：", ee.Message.ToString());
                }
            }

            string output = JsonConvert.SerializeObject(docDataList);

            writeJson(savejsontxtfile, output);

            Console.WriteLine("end");

            Console.WriteLine(output);        

            // wordDoc.SaveAs("C:\\Users\\Administrator\\Desktop\\新建文件夹 (2)\\b\\3.doc");

            //关闭wordDoc文档
            wordDoc.Close();
            wordApp.Quit();
        }

        static List<Word.Range> getTableRangeList(Word.Document wordDoc)
        {
            //把每一个表格存入一个数组       
            List<Word.Range> TablesRanges = new List<Word.Range>();
            for (int itable = 1; itable <= wordDoc.Tables.Count; itable++)
            {
                Word.Range TRange = wordDoc.Tables[itable].Range;
                TablesRanges.Add(TRange);
            }
            return TablesRanges;
        }

        static List<String[,]> getTableDataList(Word.Document wordDoc)
        {
            //表格中的数据
            String[,] tableData;
            List<String[,]> tableDataList = new List<String[,]>();
            //开始循环
            string cell_string;
            int tmp_i;
            int tmp_j;
            for (int itable = 1; itable <= wordDoc.Tables.Count; itable++)
            {
                //
                tableData = new String[wordDoc.Tables[itable].Rows.Count, wordDoc.Tables[itable].Columns.Count];
                tmp_i = 0;

                //把表格中的单元格式内容遍历出来
                foreach (Word.Row row in wordDoc.Tables[itable].Rows)
                {
                    tmp_j = 0;

                    foreach (Word.Cell cell in row.Cells)
                    {
                        cell_string = cell.Range.Text.Trim();
                        if (cell_string == "\a")
                        {
                            continue;
                        }
                        cell_string = cell_string.Replace('\a', '\0');
                        cell_string = cell_string.Replace('\r', '\0');
                        cell_string = cell_string.Trim();

                        tableData[tmp_i, tmp_j] = cell_string;

                        tmp_j++;
                    }
                    tmp_i++;
                }
                tableDataList.Add(tableData);
            }
            return tableDataList;
        }

        static void writeJson(string jsonFname, string jsonString)
        {
            FileStream fs;
            if (!System.IO.File.Exists(jsonFname))
            {
                fs = new FileStream(jsonFname, FileMode.Create, FileAccess.Write);//创建写入文件 
            }
            else
            {
                fs = new FileStream(jsonFname, FileMode.Open, FileAccess.Write);
            }

            StreamWriter sr = new StreamWriter(fs);
            sr.WriteLine(jsonString);//开始写入值
            sr.Close();
            fs.Close();
        }

    }
}
