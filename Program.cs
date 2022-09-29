using System;
using System.IO;
using System.Collections.Generic;
using OfficeOpenXml;
using System.Drawing;
using OfficeOpenXml.Drawing;

decimal RoundNearstCent(decimal input)
{
    return Math.Ceiling(input * 20) / 20;
}

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

        var lpmMargin = new decimal(0.3);
        var repMargin = new decimal(0.05);
        //var l1QtyMargin = new decimal(0.15);
        var l2QtyMargin = new decimal(0.02);
        var l3QtyMargin = new decimal(0.07);
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


                    //decimal cost = decimal.Parse(item.Cost.Replace("$", ""));
                    //decimal cost = new decimal(100);
                    decimal price = RoundNearstCent((costDec * totalMargin) + costDec);
                    decimal repCommision = price * repMargin;

                    string Price = (price).ToString("$0.00");
                    //string PriceL1 = (RoundNearstCent((price * l1QtyMargin) + price)).ToString("$0.00");
                    string PriceL2 = (RoundNearstCent(price - (price * l2QtyMargin))).ToString("$0.00");
                    string PriceL3 = (RoundNearstCent(price - (price * l3QtyMargin))).ToString("$0.00");
                    string RepCommision = (repCommision).ToString("$0.00");
                    string Profit = (price - (costDec + repCommision)).ToString("$0.00");



                    decimal repPriceInt = costDec * (totalMargin + 1);


                    decimal finalPrice = Math.Ceiling(repPriceInt * 20) / 20;
                    //Console.WriteLine($"Rep Price: {finalPrice.ToString("#.##")}");
                    
                    //decimal lpmProfit = (finalPrice - costDec) - repCommision;

                    ws.Cells[row, 4].Value = Price;

                    ws.Cells[row, 5].Value = $"10-25: {PriceL2}";
                    ws.Cells[row, 6].Value = $"26+: {PriceL3}";
                    //ws.Cells[row, 7].Value = lpmProfit.ToString("$0.00");

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


