using Aspose.Cells;
using Madam_Nanomaterial_Extract.Const;
using Madam_Nanomaterial_Extract.Utils;
using System;
using System.Collections.Generic;
using System.Text;

namespace Madam_Nanomaterial_Extract.Model
{
    class NanomaterialsModel
    {
        public string Part_Designation { get; set; }
        public string Supplier_Identification { get; set; }
        public string Identification_Number { get; set; }
        public string  Drawing_Number { get; set; }
        public string Information_Available { get; set; }
        public string  Comment { get; set; }
        public string Material { get; set; }
        public string  Cas_Number { get; set; }
        public string Size { get; set; }
        public string Function_of_the_Nanomaterials { get; set; }
        public string Integration_into_the_Article { get; set; }
        public string Form_of_Nanomaterials { get; set; }
        public string Possible_release_in_Use { get; set; }
        public string Weight_Percentage { get; set; }
        public string Weight_of_the_Article { get; set; }


        public NanomaterialsModel(Worksheet sheet, int j)
        {
            Part_Designation = UtilsCell.getCellStringValue(sheet, j + Consts.Part_Designation, 2);
            Supplier_Identification = UtilsCell.getCellStringValue(sheet, j + Consts.Supplier_Identification, 2);
            Identification_Number = UtilsCell.getCellStringValue(sheet, j + Consts.Identification_Number, 2);
            Drawing_Number = UtilsCell.getCellStringValue(sheet, j + Consts.Drawing_Number, 2);
            Information_Available = UtilsCell.getCellStringValue(sheet, j + Consts.Information_Available, 2);
            Comment = UtilsCell.getCellStringValue(sheet, j + Consts.Comment, 2);
            Material = UtilsCell.getCellStringValue(sheet, j + Consts.Material, 2);
            Cas_Number = UtilsCell.getCellStringValue(sheet, j + Consts.Cas_Number, 2);
            Size = UtilsCell.getCellStringValue(sheet, j + Consts.Size, 2);
            Function_of_the_Nanomaterials = UtilsCell.getCellStringValue(sheet, j + Consts.Function_of_the_Nanomaterials, 2);
            Integration_into_the_Article = UtilsCell.getCellStringValue(sheet, j + Consts.Integration_into_the_Article, 2);
            Form_of_Nanomaterials = UtilsCell.getCellStringValue(sheet, j + Consts.Form_of_Nanomaterials, 2);
            Possible_release_in_Use = UtilsCell.getCellStringValue(sheet, j + Consts.Possible_release_in_Use, 2);
            Weight_Percentage = UtilsCell.getCellStringValue(sheet, j + Consts.Weight_Percentage, 2);
            Weight_of_the_Article = UtilsCell.getCellStringValue(sheet, j + Consts.Weight_of_the_Article, 2);
        }

    }
}
