using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

using ExcelMerge.Model;

namespace ExcelMerge
{
    public partial class FormMain : Form
    {
        #region Properties
        private List<OutputData> _resultOutput;
        #endregion

        #region FormMain
        public FormMain()
        {
            InitializeComponent();
        }
        #endregion

        #region Actions
        private void UpdateUI()
        {
            bool enable = !String.IsNullOrWhiteSpace(textBoxInput1.Text) && !String.IsNullOrWhiteSpace(textBoxInput2.Text);

            buttonProcess.Enabled = enable;
            buttonSave.Enabled = enable;
        }

        private void buttonOpenFile1_Click(object sender, EventArgs e)
        {
            textBoxInput1.Text = GetExcelFileName("Odaberite prvu Excel datoteku");

            UpdateUI();
        }

        private void buttonOpenFile2_Click(object sender, EventArgs e)
        {
            textBoxInput2.Text = GetExcelFileName("Odaberite drugu Excel datoteku");

            UpdateUI();
        }
        
        private string GetExcelFileName(string title)
        {
            using (var ofd = new OpenFileDialog())
            {
                ofd.Title = title;
                ofd.Filter = "Excel datoteke|*.xlsx;*.xls";

                DialogResult res = ofd.ShowDialog();
                if (res == DialogResult.OK)
                {
                     return ofd.FileName;
                }
            }

            return String.Empty;
        }

        private void buttonProcess_Click(object sender, EventArgs e)
        {
            try
            {
                var inputs1 = ProcessInputFile(textBoxInput1.Text);
                var inputs2 = ProcessInputFile(textBoxInput2.Text);

                _resultOutput = new List<OutputData>();
                foreach (var input1 in inputs1)
                {
                    if (inputs2.ContainsKey(input1.Key))
                    {
                        var kolicina1 = input1.Value.Kolicina;
                        var kolicina2 = inputs2[input1.Key].Kolicina;
                        var stanje = kolicina2 - kolicina1;
                        _resultOutput.Add(new OutputData()
                        { 
                            Sifra = input1.Value.Sifra,
                            Naziv = input1.Value.Naziv,
                            Kolicina1 = kolicina1,
                            Kolicina2 = kolicina2,
                            Stanje = kolicina2 - kolicina1
                        });
                    }
                }

                dataGridViewResult.DataSource = _resultOutput;
            }
            catch (Exception excp)
            {
                MessageBox.Show(excp.Message, "Greška!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #region Input file processing
        private Dictionary<string, InputData> ProcessInputFile(string filename)
        {
            var input1 = new Dictionary<string, InputData>();
            using (var strm = new System.IO.FileStream(filename, System.IO.FileMode.Open))
            {
                IWorkbook workbook = new XSSFWorkbook(strm);
                ISheet sheetSrc = workbook.GetSheetAt(0);
                if (sheetSrc == null)
                {
                    MessageBox.Show("Neuspješno dohvaćanje lista iz tablice", "Greška!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    ValidateFirstRow(sheetSrc);

                    for (int rownum = 1; rownum < sheetSrc.PhysicalNumberOfRows; ++rownum)
                    {
                        var row = (XSSFRow)sheetSrc.GetRow(rownum);
                        if (row != null)
                        {
                            InputData data = ProcessInputRow(row);
                            if (data != null)
                            {
                                input1.Add(data.Sifra, data);
                            }
                        }
                    }
                }
            }
            return input1;
        }

        private InputData ProcessInputRow(XSSFRow row)
        {
            if (row.LastCellNum == 3)
            {
                string sifra = String.Empty;
                if (row.Cells[0].CellType == CellType.String)
                {
                    sifra = row.Cells[0].StringCellValue;
                }
                else if (row.Cells[0].CellType == CellType.Numeric)
                {
                    sifra = row.Cells[0].ToString();
                }
                else
                {
                    sifra = "GREŠKA!";
                }
                string naziv = row.Cells[1].StringCellValue;
                double kolicina = row.Cells[2].NumericCellValue;

                InputData data = new InputData()
                {
                    Sifra = sifra,
                    Naziv = naziv,
                    Kolicina = kolicina
                };

                return data;
            }

            return null;
        }

        private static void ValidateFirstRow(ISheet sheetSrc)
        {
            IRow firstRow = sheetSrc.GetRow(0);

            if (firstRow == null)
            {
                throw new Exception("Prvi list u tablici nema redaka.");
            }
            if (firstRow.Cells[0].StringCellValue != "Sifra")
            {
                throw new Exception(String.Format("Prvi stupac prve datoteke mora sadržavati šifre, a ne {0}", firstRow.Cells[0].StringCellValue));
            }
            if (firstRow.Cells[1].StringCellValue != "Naziv")
            {
                throw new Exception(String.Format("Drugi stupac prve datoteke mora sadržavati nazive, a ne {0}", firstRow.Cells[1].StringCellValue));
            }
            if (firstRow.Cells[2].StringCellValue != "Kolicina")
            {
                throw new Exception(String.Format("Treći stupac prve datoteke mora sadržavati količine, a ne {0}", firstRow.Cells[2].StringCellValue));
            }
        }
        #endregion

        private void buttonSave_Click(object sender, EventArgs e)
        {
            try
            {
                var workbook = CreateExceWorkBookFromResult();
                if (workbook == null)
                {
                    throw new Exception("Nema podataka za spremiti!");
                }

                using (var sfd = new SaveFileDialog())
                {
                    sfd.Title = "Spremi kao...";
                    sfd.Filter = "Excel datoteke|*.xlsx;*.xls";

                    DialogResult res = sfd.ShowDialog();
                    if (res == DialogResult.OK)
                    {
                        using (var strm = new System.IO.FileStream(sfd.FileName, System.IO.FileMode.Create))
                        {
                            workbook.Write(strm);
                        }
                    }
                }
            }
            catch (Exception excp)
            {
                MessageBox.Show(excp.Message, "Greška!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private IWorkbook CreateExceWorkBookFromResult()
        {
            IWorkbook workbook = null;
            if (_resultOutput != null)
            {
                workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("Rezultat");
                
                IRow headerRow = sheet.CreateRow(0);
                headerRow.CreateCell(0).SetCellValue("Sifra");
                headerRow.CreateCell(1).SetCellValue("Naziv");
                headerRow.CreateCell(2).SetCellValue("Kolicina1");
                headerRow.CreateCell(3).SetCellValue("Kolicina2");
                headerRow.CreateCell(4).SetCellValue("Stanje"); 

                int rownum = 1;
                foreach (var item in _resultOutput)
                {
                    IRow row = sheet.CreateRow(rownum++);

                    row.CreateCell(0).SetCellValue(item.Sifra);
                    row.CreateCell(1).SetCellValue(item.Naziv);
                    row.CreateCell(2).SetCellValue(item.Kolicina1);
                    row.CreateCell(3).SetCellValue(item.Kolicina2);
                    row.CreateCell(4).SetCellValue(item.Stanje); 
                }
            }

            return workbook;
        }

        #endregion

    }
}
