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
            textBoxInput1.Text = GetExcelFileName("Odaberite Excel datoteku s potrebama");

            UpdateUI();
        }

        private void buttonOpenFile2_Click(object sender, EventArgs e)
        {
            textBoxInput2.Text = GetExcelFileName("Odaberite Excel datoteku s trenutnim stanjem skladišta");

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
                var potrebe = ProcessInputFile(textBoxInput1.Text);
                var skladiste = ProcessInputFile(textBoxInput2.Text);

                _resultOutput = new List<OutputData>();
                foreach (var potreba in potrebe)
                {
                    var potrebnaKolicina = potreba.Value.Kolicina;
                    if (skladiste.ContainsKey(potreba.Key))
                    {
                        var kolicinaNaSkladistu = skladiste[potreba.Key].Kolicina;
                        var stanjeNaSkladistu = kolicinaNaSkladistu - potrebnaKolicina;

                        if (stanjeNaSkladistu < 0)
                        {
                            _resultOutput.Add(new OutputData()
                            {
                                Sifra = potreba.Value.Sifra,
                                Naziv = potreba.Value.Naziv,
                                Potreba = potrebnaKolicina,
                                Skladiste = kolicinaNaSkladistu,
                                ZaNaručiti = -stanjeNaSkladistu
                            });
                        }
                    }
                    else                  
                    {
                        _resultOutput.Add(new OutputData()
                        {
                            Sifra = potreba.Value.Sifra,
                            Naziv = String.Format("NEMA ŠIFRE NA SKLADIŠTU ZA [{0}]", potreba.Value.Naziv),
                            Potreba = potrebnaKolicina,
                            Skladiste = 0,
                            ZaNaručiti = potrebnaKolicina
                        });
                    }
                }

                dataGridViewResult.DataSource = _resultOutput;
                groupBoxResult.Text = String.Format("Generirano redaka: {0} ", _resultOutput.Count);
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
                var redStyle = workbook.CreateCellStyle();
                redStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Red.Index;
                redStyle.FillPattern = FillPattern.SolidForeground;

                ISheet sheet = workbook.CreateSheet("Rezultat");
                
                IRow headerRow = sheet.CreateRow(0);
                headerRow.CreateCell(0).SetCellValue("Sifra");
                headerRow.CreateCell(1).SetCellValue("Naziv");
                headerRow.CreateCell(2).SetCellValue("Potreba");
                headerRow.CreateCell(3).SetCellValue("Skladiste");
                headerRow.CreateCell(4).SetCellValue("Za naručiti"); 

                int rownum = 1;
                foreach (var item in _resultOutput)
                {
                    IRow row = sheet.CreateRow(rownum++);

                    row.CreateCell(0).SetCellValue(item.Sifra);
                    ICell nazivCell = row.CreateCell(1);
                    if (item.Naziv.StartsWith("NEMA ŠIFRE NA SKLADIŠTU"))
                    {
                        nazivCell.CellStyle = redStyle;
                    }
                    nazivCell.SetCellValue(item.Naziv);

                    row.CreateCell(2).SetCellValue(item.Potreba);
                    row.CreateCell(3).SetCellValue(item.Skladiste);
                    row.CreateCell(4).SetCellValue(item.ZaNaručiti); 
                }

                sheet.SetColumnWidth(0, 3500);
                sheet.SetColumnWidth(1, 18000);
                sheet.SetColumnWidth(2, 3000);
                sheet.SetColumnWidth(3, 3000);
                sheet.SetColumnWidth(4, 3000);
            }

            return workbook;
        }

        #endregion

        private void dataGridViewResult_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.ColumnIndex == 1)
            {
                string naziv = e.Value as string;
                if (naziv != null)
                {
                    if (naziv.StartsWith("NEMA ŠIFRE NA SKLADIŠTU"))
                    {
                        e.CellStyle.BackColor = Color.Red;
                    }
                }
            }
        }

    }
}
