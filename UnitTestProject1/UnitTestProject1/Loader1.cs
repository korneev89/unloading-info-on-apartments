using System;
using System.Threading;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

using System.Data;
using System.Text;
using GemBox.Spreadsheet;
using System.Linq;

namespace csharp_example
{
	[TestFixture]
	public class Loader1 : TestBase
	{
		[SetUp]
		public void Start()
		{
			driver = new ChromeDriver();
			wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
		}

		[Test]
		public void LoadInfo()
		{
			
			driver.Url = "http://lubercysamolet.ru/kvartiry/";
			driver.Manage().Window.Maximize();

			//select 2018 year only
			driver.FindElement(By.CssSelector("div.b-switcher__triggers > ul > li:nth-child(3)")).Click();
			// check "в продаже"
			driver.FindElements(By.CssSelector("div.b-search-filter__other label"))[1].Click();
			// показать n квартир
			driver.FindElement(By.CssSelector("a.b-search-filter__submit-btn.b-btn")).Click();
			// показать ещё
			var moreButton = 
				driver.FindElement(By.CssSelector("#building-11 > div > div > div > div > div > button"));

			Thread.Sleep(1500);
			while (IsElementVisible(driver, By.CssSelector("#building-11 > div > div > div > div > div > button")))
			{
				//wait.Until(ExpectedConditions.ElementToBeClickable(moreButton));
				moreButton.Click();
				Thread.Sleep(500);
			}
			//Thread.Sleep(30000);

			SpreadsheetInfo.SetLicense("EIKU-U5LX-6MSF-Z84S");
			ExcelFile ef = new ExcelFile();
			ExcelWorksheet ws = ef.Worksheets.Add("выгрузка");

			DataTable dt = new DataTable();

			// add columns
			var c1 = dt.Columns.Add("Ссылка", typeof(string));
			var c2 = dt.Columns.Add("Очередь", typeof(string));
			var c3 = dt.Columns.Add("Корпус", typeof(int));
			var c4 = dt.Columns.Add("Номер", typeof(int));
			var c5 = dt.Columns.Add("Комнат", typeof(int));
			var c6 = dt.Columns.Add("Общая", typeof(double));
			var c7 = dt.Columns.Add("Жилая", typeof(double));
			var c8 = dt.Columns.Add("Этаж", typeof(int));
			var c9 = dt.Columns.Add("Оплата", typeof(long));
			var c10 = dt.Columns.Add("Статус", typeof(string));
			var c11 = dt.Columns.Add("Отделка", typeof(string));


			var rows = driver.FindElements(By.CssSelector("tr.j-building-tr-link"));
			foreach (IWebElement row in rows)
			{
				var v1 = row.FindElements(By.CssSelector("td"))[1].GetAttribute("textContent").
				Replace("\r\n","").Replace("  "," ").Replace("      ", "").Replace("    ", "");

				var v2 = row.FindElements(By.CssSelector("td"))[2].GetAttribute("textContent");
				var r2 = Int32.Parse(new String(v2.Where(Char.IsDigit).ToArray()));

				var v3 = row.FindElements(By.CssSelector("td"))[3].GetAttribute("textContent");
				var r3 = Int32.Parse(new String(v3.Where(Char.IsDigit).ToArray()));
				var link = row.FindElement(By.CssSelector("td a")).GetAttribute("href");

				var v4 = row.FindElements(By.CssSelector("td"))[4].GetAttribute("textContent");
				var r4 = Int32.Parse(new String(v4.Where(Char.IsDigit).ToArray()));

				var v5 = row.FindElements(By.CssSelector("td"))[5].GetAttribute("textContent")
					.Replace("Общая", "")
					.Replace("м2", "")
					.Replace("\r\n", "")
					.Replace(Convert.ToChar(160).ToString(), "")
					.Replace(" ", "")
					.Replace(".",",");
				var r5 = Double.Parse(v5);

				var v6 = row.FindElements(By.CssSelector("td"))[6].GetAttribute("textContent")
					.Replace("Жилая", "")
					.Replace("м2", "")
					.Replace("\r\n", "")
					.Replace(Convert.ToChar(160).ToString(), "")
					.Replace(" ", "")
					.Replace(".", ",");
				var r6 = Double.Parse(v6);

				var v7 = row.FindElements(By.CssSelector("td"))[7].GetAttribute("textContent");
				var r7 = Int32.Parse(new String(v7.Where(Char.IsDigit).ToArray()));

				var v8 = row.FindElements(By.CssSelector("td"))[8].FindElement(By.CssSelector("span.b-building__price")).GetAttribute("textContent");
				var r8 = Int64.Parse(new String(v8.Where(Char.IsDigit).ToArray()));

				var v9 = row.FindElements(By.CssSelector("td"))[9].FindElement(By.CssSelector("div")).GetAttribute("textContent").
					Replace("\r\n", "").Replace("  ", " ").Replace("        ","").Replace("      ", "");

				var v10 = row.FindElements(By.CssSelector("td"))[9].FindElements(By.CssSelector("div.b-decoration"));
				string r10;
				if (v10.Count > 0)
				{
					r10 = v10[0].GetAttribute("textContent").Replace("\r\n", "").Replace("  ", " ").Replace("                                        ", "");
				}
				else
				{
					r10 = "С отделкой";
				}

				dt.Rows.Add(
					link,
					v1, 
					r2,
					r3,
					r4,
					r5,
					r6,
					r7,
					r8,
					v9,
					r10);
		
			}

			// add cell
			// ws.Cells[0, 0].Value = "DataTable insert example:";

			// Insert DataTable into an Excel worksheet.
			ws.InsertDataTable(dt,
				new InsertDataTableOptions()
				{
					ColumnHeaders = true,
					StartRow = 0
				});
			// Autofit columns and some print options (for better look when exporting to pdf, xps and printing).
			var columnCount = ws.CalculateMaxUsedColumns();
			for (int i = 0; i < columnCount; i++)
				ws.Columns[i].AutoFit();

			var date = DateTime.Now.ToString("yyyy.MM.dd_HH-mm-ss");
			var fileName = String.Concat(@"D:/выгрузка_",date,".xlsx");
			ef.Save(fileName);
		}
	}
}
