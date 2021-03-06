﻿using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;

namespace csharp_example
{
	[TestFixture]
	public class TestBase
	{
		public IWebDriver driver;
		public WebDriverWait wait;

		[TearDown]
		public void Stop()
		{
			driver.Quit();
			driver = null;
		}

		public void LoginAsAdmin()
		{
			driver.Url = "http://localhost:8080/litecart/admin/login.php";
			driver.FindElement(By.Name("username")).SendKeys("admin");
			driver.FindElement(By.Name("password")).SendKeys("admin");
			driver.FindElement(By.Name("login")).Click();
		}

		public Boolean IsElementPresent(IWebDriver driver, By locator)
		{
			try
			{
				wait = new WebDriverWait(driver, TimeSpan.FromSeconds(1));
				return driver.FindElements(locator).Count > 0;
			}
			finally
			{
				wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
			}
		}

		public Boolean IsElementVisible(IWebDriver driver, By locator)
		{
			try
			{
				wait = new WebDriverWait(driver, TimeSpan.FromSeconds(1));
				wait.Until(ExpectedConditions.ElementIsVisible(locator));
				return true;
			}
			catch
			{
				return false;
			}
			finally
			{
				wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
			}
		}

		public static Func<IWebDriver, string> ThereIsWindowOtherThan(IEnumerable<string> oldWindows)
		{
			string GetNewWindow(IWebDriver driver)
			{
				return driver.WindowHandles.Except(oldWindows).ToList().Single();
			}

			return GetNewWindow;
		}
	}
}