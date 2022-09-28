using System;
using System.IO;
using System.Collections.Generic;
using OfficeOpenXml;
using System.Drawing;
using OfficeOpenXml.Drawing;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
using (var excelPack = new ExcelPackage())
{
    //Load excel stream
    using (var stream = File.OpenRead("../../../BestLighting/2022-05-02.xlsx"))
    {
        excelPack.Load(stream);


        //Lets Deal with first worksheet.(You may iterate here if dealing with multiple sheets)
        var ws = excelPack.Workbook.Worksheets[0];


        //Console.WriteLine(ws.Cells[1,1].Value);

        var start = ws.Dimension.Start;
        var end = ws.Dimension.End;

        var lpmMargin = new decimal(0.35);
        var repMargin = new decimal(0.05);
        var totalMargin = lpmMargin + repMargin;
        Console.WriteLine(totalMargin);
        //var inProduct = false;

        for (int row = start.Row; row <= end.Row; row++)
        {

            var cell1 = ws.Cells[row, 1].Text.Trim();
            var cell2 = ws.Cells[row, 2].Text.Trim();
            var cell3 = ws.Cells[row, 3].Text.Trim();
            var cell4 = ws.Cells[row, 4].Text.Trim();

            //if (!inProduct && (cell1.StartsWith("Model") || cell1.StartsWith("Base Product")))
            //{
            //    inProduct = true;
            //}
            //else if (inProduct && cell1.StartsWith("Sample Part Nbr:"))
            //{
            //    inProduct = false;
            //}


            //Console.WriteLine(inProduct);

            //Console.WriteLine($"'{cell1}'");
            //Console.WriteLine(cell2);
            //Console.WriteLine(cell3);
            //Console.WriteLine(cell4);        



            if (cell4.StartsWith("$"))
            {
                var cost = cell4.Trim().Replace("$", "");                
                decimal costDec = decimal.Parse(cost);
                if(costDec > 0)
                {
                    decimal repPriceInt = costDec * (totalMargin + 1);


                    decimal finalPrice = Math.Ceiling(repPriceInt * 20) / 20;
                    //Console.WriteLine($"Rep Price: {finalPrice.ToString("#.##")}");

                    decimal repCommision = repMargin * finalPrice;
                    decimal lpmProfit = (finalPrice - costDec) - repCommision;

                    ws.Cells[row, 4].Value = finalPrice.ToString("$0.00");

                    ws.Cells[row, 5].Value = costDec.ToString("$0.00");
                    ws.Cells[row, 6].Value = repCommision.ToString("$0.00");
                    ws.Cells[row, 7].Value = lpmProfit.ToString("$0.00");

                }


            }

            //Console.WriteLine("==================================================================");
        }

        string path = $@"..\..\..\margin{(totalMargin*100).ToString("0")}.xlsx";
        Stream outStream = File.Create(path);
        excelPack.SaveAs(outStream);
        stream.Close();

    }
    //var she}et = ws.Workbook.Worksheets[sheetname];
    //var pic = ws.Drawings["image2.jpeg"] as ExcelPicture;        

}
