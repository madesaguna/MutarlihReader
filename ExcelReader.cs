using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml;

namespace ImportPemutakhiran
{
    class ExcelReader
    {
        FileInfo fileName;
        static string TABLE_NAME = "pemutakhiran";
        static string[] HEADER =  {
                "no", "dpid", "no_kk", "nik", "nama_lengkap",
                "tempat_lahir", "tanggal_lahir", "status_kawin",
                "jenis_kelamin", "alamat", "rt", "rw", "disabilitas",
                "keterangan", "tps", "sheetname", "nama_kelurahan",
                "nama_kecamatan", "nama_kabupaten"
            };
        Dictionary<string, int> maxColumns = new Dictionary<string, int> {
                { "SARING", 15 }, {"UBAH", 15}, { "BARU", 14 }, { "PINDAH", 15 }
            };
        public ExcelReader(string fileName)
        {
            this.fileName = new FileInfo(fileName);
        }

        private string generateSQL(Dictionary<string, string> dict)
        {
            var sql = new StringBuilder();
            sql.Append($"INSERT INTO `{TABLE_NAME}` (");
            sql.Append(String.Join(",", HEADER.Select(z => $"`{z}`").ToArray()));
            sql.Append(") VALUES (");
            sql.Append(String.Join(",", dict.Select(z => $"'{z.Value.Replace("'", "\'")}'").ToArray()));
            sql.Append(");");
            return sql.ToString();
        }

        public void Execute()
        {
            foreach (var sheet in maxColumns.Keys)
            {
                ReadSheet(sheet).ForEach(_dict =>
                {
                    Console.WriteLine(generateSQL(_dict));
                });
            }
        }
        /* https://stackoverflow.com/questions/7252186/switch-case-on-type-c-sharp */
        private string GetValue(ExcelRange s)
        {
            var _v = s.Value;
            string res;
            if (_v == null)
            {
                res = "";
            }
            else if (_v.GetType() == typeof(double))
            {
                res = _v.ToString();
            }
            else
            {
                res = _v.ToString();
            }
            return res;
        }

        private Dictionary<string, string> ReadColumn(ExcelWorksheet sheet, string sheetName, int startRead)
        {
            Dictionary<string, string> _dict = new Dictionary<string, string>();
            int y;
            for (int i = 0; i < maxColumns[sheetName]; i++)
            {
                y = i + 1;
                if (string.Equals(sheetName, "baru", StringComparison.OrdinalIgnoreCase))
                {
                    if (y == 1)
                    {
                        _dict.Add(HEADER[i], GetValue(sheet.Cells[startRead, y]));
                        
                    }
                    else if (y == 2)
                    {
                        _dict.Add(HEADER[i], "0");
                    }
                    else
                    {
                        _dict.Add(HEADER[y], GetValue(sheet.Cells[startRead, y]));
                    }
                }
                else
                {
                    _dict.Add(HEADER[i], GetValue(sheet.Cells[startRead, y]));
                }

            }
            if (string.Equals(sheetName, "pindah", StringComparison.OrdinalIgnoreCase))
            {
                _dict["tps"] = GetValue(sheet.Cells[startRead, 16]);
                _dict["keterangan"] = "P";
            }

            if (string.Equals(sheetName, "baru", StringComparison.OrdinalIgnoreCase))
            {
                _dict["keterangan"] = "B";
            }
            _dict.Add("sheetname", sheetName.ToLower());
            return _dict;
        }

        private List<Dictionary<string, string>> ReadSheet(string sheetName)
        {
            var elist = new List<Dictionary<string, string>>();
            int indexFullName = (string.Equals(sheetName, "baru", StringComparison.OrdinalIgnoreCase)) ? 4 : 5;
            using (var p = new ExcelPackage(this.fileName))
            {
                var ws = p.Workbook.Worksheets[sheetName];
                int startRead = 11;
                while (true)
                {
                    if (ws.Cells[startRead, indexFullName].Value is null)
                    {
                        break;
                    }
                    var cell = ws.Cells[startRead, 1];
                    elist.Add(ReadColumn(ws, sheetName, startRead));
                    startRead++;

                }
                return elist;

                //The style object is used to access most cells formatting and styles.
                //ws.Cells[3, 1].Style.Font.Bold = true;
                //p.Save();
            }
        }
    }
}