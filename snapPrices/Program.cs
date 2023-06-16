
// See https://aka.ms/new-console-template for more information
using HtmlAgilityPack;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;


using System.Net;

using IronXL;
using OpenQA.Selenium.Interactions;

Dictionary<string, Dictionary<string, string>> data = new Dictionary<string, Dictionary<string, string>>();
List<string> important = new List<string>();
important.Add("گوجه");
important.Add("سیب");
important.Add("سیب زمینی");
important.Add("خیار");
important.Add("آلو");
important.Add("آلبالو");
important.Add("شلیل");
important.Add("هندوانه");
important.Add("پیاز");
important.Add("هلو");
important.Add("زردآلو");
important.Add("کدو");
important.Add("کلم");
important.Add("بادمجان");
important.Add("فلفل");
important.Add("سیر");
important.Add("موز");
important.Add("آناناس");
important.Add("انبه");
important.Add("آواکاو");
important.Add("پرتقال");
important.Add("خرما");
important.Add("هویج");
important.Add("گیلاس");
important.Add("قارچ");
important.Add("کاهو");
important.Add("طالبی");
important.Add("خربزه");
important.Add("انگور");
important.Add("لیمو");
important.Add("بامیه");
List<string> One = new List<string>();
One.Add("هلو");

One.Add("موز");
One.Add("شلیل");
One.Add("زردآلو");
One.Add("انبه");
One.Add("آناناس");
One.Add("بادمجان");
One.Add("هویج");
One.Add("قارچ");
One.Add("گیلاس");
One.Add("کدو");
One.Add("آلبالو");
One.Add("هندوانه");
One.Add("سیر");
One.Add("طالبی");
One.Add("خربزه");

List<string> two = new List<string>();
two.Add("کلم");
two.Add("سیب");
two.Add("خیار");
two.Add("گوجه");
two.Add("فلفل");
two.Add("آلو");
two.Add("کاهو");
two.Add("انگور");


List<string> html = new List<string>();
//html.Add(@"https://snappfood.ir/grocery/menu/%D8%B3%D9%88%D9%BE%D8%B1_%D9%85%DB%8C%D9%88%D9%87_%D9%88%DB%8C%D8%AA%D8%A7%D9%85%DB%8C%D9%86_30-r-3224j2");
//html.Add(@"https://snappfood.ir/grocery/menu/%D8%B3%D9%88%D9%BE%D8%B1_%D9%85%DB%8C%D9%88%D9%87_%D8%AF%D8%B3%D8%AA%DA%86%DB%8C%D9%86-r-3d71gx");
//html.Add(@"https://snappfood.ir/grocery/menu/%D8%B3%D9%88%D9%BE%D8%B1_%D9%85%DB%8C%D9%88%D9%87_%D8%B9%D8%A7%D8%B1%D9%81-r-0n6yn2");
//html.Add(@"https://snappfood.ir/grocery/menu/%D8%B3%D9%88%D9%BE%D8%B1_%D9%85%DB%8C%D9%88%D9%87_%DA%A9%DB%8C%D9%85%DB%8C%D8%A7_2-r-0gvq4r");
//html.Add(@"https://snappfood.ir/grocery/menu/%D8%B3%D9%88%D9%BE%D8%B1_%D9%85%DB%8C%D9%88%D9%87_%D8%A7%D8%AD%D9%85%D8%AF-r-0gwrex");
//html.Add(@"https://snappfood.ir/grocery/menu/%D9%85%DB%8C%D9%88%D9%87_%D8%AA%D9%88%DA%A9%D9%84%DB%8C-r-31n5dy");
//html.Add(@"https://snappfood.ir/grocery/menu/%D8%B3%D9%88%D9%BE%D8%B1_%D9%85%DB%8C%D9%88%D9%87_%D8%A8%D8%B1%D8%A7%D8%AF%D8%B1%D8%A7%D9%86_%D8%B1%D8%B3%D8%AA%DA%AF%D8%A7%D8%B1-r-p6ygdn");
//html.Add(@"https://snappfood.ir/grocery/menu/%D8%B3%D9%88%D9%BE%D8%B1_%D9%85%DB%8C%D9%88%D9%87_%D8%B3%D8%B9%D8%AF%D8%A2%D8%A8%D8%A7%D8%AF-r-p6y2xn");
//html.Add(@"https://snappfood.ir/grocery/menu/%D8%B3%D9%88%D9%BE%D8%B1_%D9%85%DB%8C%D9%88%D9%87_%D8%B5%D8%A7%D8%AF%D9%82%DB%8C%D8%A7%D9%86-r-0g7mon");
//// Default file format is XLSX, we can override it using CreatingOptions

html.Add(@"https://snappfood.ir/grocery/menu/%D8%B3%D9%88%D9%BE%D8%B1_%D9%85%DB%8C%D9%88%D9%87_%D8%B9%D8%A7%D8%B1%D9%81-r-0n6yn2");
html.Add(@"https://snappfood.ir/grocery/menu/%D8%B3%D9%88%D9%BE%D8%B1_%D9%85%DB%8C%D9%88%D9%87_%D8%AF%D8%B3%D8%AA%DA%86%DB%8C%D9%86-r-3d71gx");
//html.Add(@"https://snappfood.ir/grocery/menu/%D8%B3%D8%A8%D8%B2%DB%8C%D8%AC%D8%A7%D8%AA_%D9%88_%D9%85%D8%AD%D8%B5%D9%88%D9%84%D8%A7%D8%AA_%D8%A7%D8%B1%DA%AF%D8%A7%D9%86%DB%8C%DA%A9_%D9%86%D9%88%D8%A8%D8%B1-r-p8gngm");
//html.Add(@"https://snappfood.ir/grocery/menu/%D8%B3%D9%88%D9%BE%D8%B1_%D9%85%DB%8C%D9%88%D9%87_%DA%A9%DB%8C%D9%85%DB%8C%D8%A7-r-0gzve2");
//html.Add(@"https://snappfood.ir/grocery/menu/%D8%B3%D9%88%D9%BE%D8%B1_%D9%85%DB%8C%D9%88%D9%87_%D8%A7%D8%AD%D9%85%D8%AF-r-0gwrex");
//html.Add(@"https://snappfood.ir/grocery/menu/%D8%B3%D9%88%D9%BE%D8%B1_%D9%85%DB%8C%D9%88%D9%87_%DA%A9%DB%8C%D9%85%DB%8C%D8%A7_2-r-0gvq4r");
//html.Add(@"https://snappfood.ir/grocery/menu/%D9%85%DB%8C%D9%88%D9%87_%D8%B3%D8%B1%D8%A7_%D8%AF%DA%A9%D8%A7%D9%86_%D9%85%D8%A7%D8%B1%DA%A9%D8%AA-r-p8goly");
//html.Add(@"https://snappfood.ir/grocery/menu/%D9%85%DB%8C%D9%88%D9%87_%D8%AA%D9%88%DA%A9%D9%84%DB%8C-r-31n5dy");
//html.Add(@"https://snappfood.ir/grocery/menu/%D8%B3%D9%88%D9%BE%D8%B1_%D9%85%DB%8C%D9%88%D9%87_%D8%B5%D8%A7%D8%AF%D9%82%DB%8C%D8%A7%D9%86-r-0g7mon");
//html.Add(@"https://snappfood.ir/grocery/menu/%D8%B3%D9%88%D9%BE%D8%B1_%D9%85%DB%8C%D9%88%D9%87_%D8%A8%D8%B1%D8%A7%D8%AF%D8%B1%D8%A7%D9%86_%D8%B1%D8%B3%D8%AA%DA%AF%D8%A7%D8%B1-r-p6ygdn");
//html.Add(@"https://snappfood.ir/grocery/menu/%D9%85%DB%8C%D9%88%D9%87_%D9%81%D8%B1%D9%88%D8%B4%DB%8C_%DA%86%D9%87%D8%A7%D8%B1%D9%81%D8%B5%D9%84-r-3xg6xy");
//html.Add(@"https://snappfood.ir/grocery/menu/%D8%B3%D9%88%D9%BE%D8%B1_%D9%85%DB%8C%D9%88%D9%87_%D8%B3%D8%B9%D8%AF%D8%A2%D8%A8%D8%A7%D8%AF-r-p6y2xn");
//html.Add(@"https://snappfood.ir/grocery/menu/%D8%B3%D9%88%D9%BE%D8%B1_%D9%85%DB%8C%D9%88%D9%87_%D8%B3%D8%B9%D8%AF%D8%A2%D8%A8%D8%A7%D8%AF-r-p6y2xn");
//html.Add(@"https://snappfood.ir/grocery/menu/%D9%87%D8%A7%DB%8C%D9%BE%D8%B1_%D9%85%DB%8C%D9%88%D9%87_%D8%AA%D9%85%D8%B4%DA%A9-r-3xg5wy");
//html.Add(@"https://snappfood.ir/grocery/menu/%D9%85%DB%8C%D9%88%D9%87_%D8%B3%D8%B1%D8%A7%DB%8C_%D8%B4%D8%A7%D9%86%D8%AF%DB%8C%D8%B2__%D8%AF%D8%A7%D9%86%D8%B4%D8%AC%D9%88_-r-pvlqmv");
//html.Add(@"https://snappfood.ir/grocery/menu/%D8%B3%D9%88%D9%BE%D8%B1_%D9%85%DB%8C%D9%88%D9%87_%D8%B3%D8%A7%D9%85%D8%A7%D9%86%DB%8C%D9%87-r-0r6glo");
//html.Add(@"https://snappfood.ir/grocery/menu/%D8%B3%D9%88%D9%BE%D8%B1_%D9%85%DB%8C%D9%88%D9%87_%D8%A2%D8%B1%D8%A7%D9%85%D8%B4-r-3k8gqk");
//html.Add(@"https://snappfood.ir/grocery/menu/%D9%87%D8%A7%DB%8C%D9%BE%D8%B1_%D9%85%DB%8C%D9%88%D9%87_%D8%A8%D8%B2%D8%B1%DA%AF_%D8%A7%D8%AA%D9%85%D8%A7-r-0ygkqy");
//html.Add(@"https://snappfood.ir/grocery/menu/%D9%85%DB%8C%D9%88%D9%87_%D8%AA%D9%84%D9%81%D9%86%DB%8C_%D8%A8%D8%A7%D8%AC%D9%86%D8%A7%D9%82_%D9%87%D8%A7-z-3kqw4k");
//html.Add(@"https://snappfood.ir/grocery/menu/%D9%85%DB%8C%D9%88%D9%87_%D9%88%D8%A7%D8%B1%D9%84%DB%8C-r-3x61x4");
//html.Add(@"https://snappfood.ir/grocery/menu/%D8%B3%D9%88%D9%BE%D8%B1_%D9%85%DB%8C%D9%88%D9%87_%D8%AF%D8%B3%D8%AA%DA%86%DB%8C%D9%86_%D8%A8%D9%87%D8%A7%D8%B1%DB%8C-r-37odrm");
//html.Add(@"https://snappfood.ir/grocery/menu/%D9%85%DB%8C%D9%88%D9%87_%D8%B3%D8%B1%D8%A7%DB%8C_%D8%A8%D9%87%D8%B4%D8%AA-r-3d9mwx");


WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
var workSheet = workBook.CreateWorkSheet("Snap Prices");

char coloumn = 'B';
for (int c = 0; c < html.Count; c++)
{


    Console.WriteLine("complete : " + ((c + 1) * 100 / html.Count).ToString() + "%");

    IWebDriver driver = new ChromeDriver();
    driver.Navigate().GoToUrl(html[c]);

    Thread.Sleep(15000);

    var element = driver.FindElement(By.ClassName("hMDcUg"));

    element.Click();



    var createPageLink = driver.FindElements(By.ClassName("jQhxGx"));
    int counter = 1;

    Actions actions = new Actions(driver);
    actions.MoveToElement(createPageLink[createPageLink.Count - 1]);
    actions.Perform();
    element = driver.FindElement(By.ClassName("kNFBOq"));
    string storeName = element.Text;

    workSheet[coloumn + "1"].StringValue = storeName;
    foreach (var item in createPageLink)
    {
      

        var productTag = item.FindElement(By.ClassName("esHHju"));
        var priceTag = item.FindElements(By.ClassName("hxREoh"));
        var weightTag = item.FindElements(By.ClassName("sFIZX"));

        string productName = productTag.Text;
        string price = "";
        double priceLong = 0;
        for (int i = 0; i < priceTag.Count; i++)
        {
            if (weightTag.Count != priceTag.Count)
            {
                break;
            }

            string[] temp = weightTag[i].Text.Split(' ');

            string t = priceTag[i].Text;
            if (t != "")
            {
                string priceTagEdited = toEnglishNumber(t.Replace(",", ""), false).Split(' ')[0];
                for (int j = 0; j < temp.Length; j++)
                {
                    temp[j] = toEnglishNumber(temp[j], false);
                }
                if (temp.Length < 2)
                {
                    break;
                }
                if (temp[1] == "بسته")
                {
                    break;
                }
                if (weightTag[i].Text.Contains("نیم"))
                {
                    break;
                }
                double priceTagEditedDouble = 0;
                if(!double.TryParse(priceTagEdited, out priceTagEditedDouble))
                {
                    break;
                }
                
                if (temp.Length == 4)
                {
                    double wDouble = 0;
                    if (!double.TryParse(temp[2], out wDouble))
                    {
                        break;
                    }
                    if (temp[3] == "گرم")
                    {
                        temp[2] = toEnglishNumber(temp[2], true);
                        priceLong = Math.Max(priceTagEditedDouble * 1000 / wDouble, priceLong);

                        price = priceLong.ToString();
                    }
                    if (temp[3] == "کیلوگرم")
                    {
                        temp[2] = toEnglishNumber(temp[2], true);
                        priceLong = Math.Max(priceTagEditedDouble / wDouble, priceLong);

                        price = priceLong.ToString();

                    }
                    if (temp[3] == "عدد")
                    {
                        priceLong = Math.Max(priceTagEditedDouble, priceLong);

                        price = priceLong.ToString();

                    }
                    if (temp[3] == "لیتر")
                    {
                        priceLong = Math.Max(priceTagEditedDouble, priceLong);
                        price = price = priceLong.ToString();
                    }
                }
                else
                {
                    double wDouble = 0;
                    if (!double.TryParse(temp[0], out wDouble))
                    {
                        break;
                    }
                    if (temp[1] == "گرم")
                    {

                        temp[0] = toEnglishNumber(temp[0], true);

                        priceLong = Math.Max(priceTagEditedDouble * 1000 / wDouble, priceLong);

                        price = priceLong.ToString();
                    }
                    if (temp[1] == "کیلوگرم")
                    {
                        temp[0] = toEnglishNumber(temp[0], true);
                        priceLong = Math.Max(priceTagEditedDouble / wDouble, priceLong);

                        price = priceLong.ToString();

                    }
                    if (temp[1] == "عدد")
                    {
                        priceLong = Math.Max(priceTagEditedDouble, priceLong);

                        price = priceLong.ToString();

                    }
                    if (temp[1] == "لیتر")
                    {
                        priceLong = Math.Max(priceTagEditedDouble, priceLong);
                        price = price = priceLong.ToString();
                    }





                }
            }
        }

        if (price != "")
        {
            string[] priceArray = price.Split(".");
            productName = productName.Replace(" درجه ۱", "");
            if (!important.Exists(e => e.Contains(productName.Split(" ")[0])))
            {
                continue;
            }
            if (productName.Contains("اقتصادی") || productName.Contains("ممتاز") || productName.Contains("درجه ۲"))
            {
                continue;
            }

            string[] tt = productName.Split(" ");
            if (tt.Length > 2)
            {
                productName = tt[0] + " " + tt[1];
            }

            if (One.Exists(e => e.Contains(productName.Split(" ")[0])))
            {
                if (tt.Length > 0)
                    productName = tt[0];
            }
            if (two.Exists(e => e.Contains(productName.Split(" ")[0])))
            {
                if (tt.Length > 1)
                    productName = tt[0] + " " + tt[1];
            }


            Boolean flagExist = false;
            for (int excelRowsCounter = 1; excelRowsCounter < workSheet.Rows.Length + 1; excelRowsCounter++)
            {
                string b = workSheet["A" + excelRowsCounter.ToString()].Value.ToString();
                if (b == productName)
                {
                    workSheet[(coloumn).ToString() + excelRowsCounter.ToString()].LongValue = long.Parse(priceArray[0]);
                    flagExist = true;
                    break;
                }
            }
            if (flagExist == false)
            {
                int L = workSheet.Rows.Length + 1;
                workSheet["A" + (L).ToString()].Value = productName;
                string col = (coloumn).ToString();
                workSheet[col + (L).ToString()].LongValue = long.Parse(priceArray[0]);
            }
        }
        counter++;

    }
    coloumn++;
    driver.Close();
    Thread.Sleep(60000);
}
workBook.SaveAs("d://example_workbook.xlsx");


string toEnglishNumber(string input, Boolean IstNummer)
{
    string EnglishNumbers = "";
    if (IstNummer)
    {

        input = new string(input.Where(c => char.IsDigit(c)).ToArray());
    }

    for (int i = 0; i < input.Length; i++)
    {
        if (Char.IsDigit(input[i]))
        {
            EnglishNumbers += char.GetNumericValue(input, i);
        }
        else
        {
            EnglishNumbers += input[i].ToString();
        }
    }
    return EnglishNumbers;
}

