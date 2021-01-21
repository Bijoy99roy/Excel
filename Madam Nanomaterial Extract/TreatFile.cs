using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using Aspose.Cells;
using Madam_Nanomaterial_Extract.Const;
using Madam_Nanomaterial_Extract.Model;
using Madam_Nanomaterial_Extract.Utils;

namespace Madam_Nanomaterial_Extract
{
    class TreatFile
    {
        Dictionary<string, NanomaterialsModel> nanomaterials;
        List<NanomaterialsModel> nanomaterialsList;
        Workbook wb;
        public TreatFile()
        {
            wb = new Workbook(@"D:\Bassetii\C#\Madam Nanomaterial Extract\Madam Nanomaterial Extract\3.2.H-DFR-Nanomaterials Report.xlsm");
            Worksheet sheet = wb.Worksheets["Raw_Data"];
            Worksheet sheet1 = wb.Worksheets[1];
           
            nanomaterials = new Dictionary<string, NanomaterialsModel>();
            nanomaterialsList = new List<NanomaterialsModel>();
            ReadExcel(sheet);
            WriteExcel(sheet1);
        }

        public void ReadExcel(Worksheet sheet)
        {
            int j = 0;
            for(int i = 0;i < 10;i++)
            {
                
                Console.WriteLine(UtilsCell.getCellStringValue(sheet, j + Consts.Part_Designation, 2).Trim().Length);
                if(UtilsCell.getCellStringValue(sheet, j + Consts.Part_Designation, 2).Trim().Length > 0)
                    nanomaterials[UtilsCell.getCellStringValue(sheet, j + Consts.Part_Designation, 2)] = new NanomaterialsModel(sheet,j); 
                j+=16;
            }
            foreach (KeyValuePair<string, NanomaterialsModel> entry in nanomaterials)
            {
                nanomaterialsList.Add(entry.Value);
                Console.WriteLine(entry.Key + " : " + entry.Value.Part_Designation);
            }

        }

        public void WriteExcel(Worksheet sheet)
        {
            int line = 9;
            for(int i = 0;i < nanomaterialsList.Count;i++)
            {
                sheet.Cells[line + i, Consts.Part_Designation].PutValue(nanomaterialsList[i].Part_Designation);
                sheet.Cells[line + i, Consts.Supplier_Identification].PutValue(nanomaterialsList[i].Supplier_Identification);
                sheet.Cells[line + i, Consts.Identification_Number].PutValue(nanomaterialsList[i].Identification_Number);
                sheet.Cells[line + i, Consts.Drawing_Number].PutValue(nanomaterialsList[i].Drawing_Number);
                sheet.Cells[line + i, Consts.Information_Available].PutValue(nanomaterialsList[i].Information_Available);
                sheet.Cells[line + i, Consts.Comment].PutValue(nanomaterialsList[i].Comment);
                sheet.Cells[line + i, Consts.Material].PutValue(nanomaterialsList[i].Material);
                sheet.Cells[line + i, Consts.Cas_Number].PutValue(nanomaterialsList[i].Cas_Number);
                sheet.Cells[line + i, Consts.Size].PutValue(nanomaterialsList[i].Size);
                sheet.Cells[line + i, Consts.Function_of_the_Nanomaterials].PutValue(nanomaterialsList[i].Function_of_the_Nanomaterials);
                sheet.Cells[line + i, Consts.Integration_into_the_Article].PutValue(nanomaterialsList[i].Integration_into_the_Article);
                sheet.Cells[line + i, Consts.Form_of_Nanomaterials].PutValue(nanomaterialsList[i].Form_of_Nanomaterials);
                sheet.Cells[line + i, Consts.Possible_release_in_Use].PutValue(nanomaterialsList[i].Possible_release_in_Use);
                sheet.Cells[line + i, Consts.Weight_Percentage].PutValue(nanomaterialsList[i].Weight_Percentage);
                sheet.Cells[line + i, Consts.Weight_of_the_Article].PutValue(nanomaterialsList[i].Weight_of_the_Article);

            }

            wb.Save(@"D:\Bassetii\C#\Madam Nanomaterial Extract\Madam Nanomaterial Extract\3.2.H-DFR-Nanomaterials Report1.xlsm");
        }
       
    }
}
