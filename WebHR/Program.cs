using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace WebHR
{
    class Program
    {
        
        static void Main(string[] args)
        {
            Execution();
        }
        
        public static void Execution()
        {
            DataTable table = ExcelLib.PopulateInCollection(ConfigurationManager.AppSettings["data"]);

            ChromeOptions options = new ChromeOptions();
            options.AddArguments(@"user-data-dir="+Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\profiles");
            properties.driver = new ChromeDriver(options);

            properties.driver.Navigate().GoToUrl("https://tps.webhr.co/hr/login/#noanchor");
            Thread.Sleep(10000);
            properties.driver.Manage().Window.Maximize();

            try
            {
            if (properties.driver.FindElement(By.XPath("//a[contains(text(), 'Sign In with Microsoft Azure')]")).Displayed)
            {
                properties.driver.FindElement(By.XPath("//a[contains(text(), 'Sign In with Microsoft Azure')]")).Click();
                properties.driver.SwitchTo().Window(properties.driver.WindowHandles[1]);
                properties.driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
                properties.driver.FindElement(By.Id("i0116")).SendKeys(ExcelLib.ReadData(1, "email"));
                properties.driver.FindElement(By.Id("idSIButton9")).Click();
                properties.driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
                properties.driver.FindElement(By.Id("i0118")).SendKeys(ExcelLib.ReadData(1, "password"));
                properties.driver.FindElement(By.XPath("//input[@value = 'Sign in' ]")).Click();
                properties.driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1000);
                properties.driver.FindElement(By.XPath("//input[@value = 'Yes']")).Submit();
                Console.WriteLine("sign in with azure done");
            }
            
            properties.driver.SwitchTo().Window(properties.driver.WindowHandles[0]);
            if (properties.driver.FindElement(By.XPath("//input[@id = 'btnAttendanceSignOut']")).Displayed)
                {
                    properties.driver.FindElement(By.XPath("//input[@id = 'btnAttendanceSignOut']")).Click();
                    Console.WriteLine("User has signed out");
                }

                else
                {
                    Console.WriteLine("Element not found");
                    //properties.driver.FindElement(By.XPath("//input[@id = 'btnAttendanceSignOut']")).Submit();
                }


            }
            catch (Exception ex)
            {
                Console.WriteLine("\nUser Already Logged-In");
                properties.driver.SwitchTo().Window(properties.driver.WindowHandles[0]);
                if (properties.driver.FindElement(By.XPath("//input[@id = 'btnAttendanceSignOut']")).Displayed)
                {
                    properties.driver.FindElement(By.XPath("//input[@id = 'btnAttendanceSignOut']")).Click();
                    Console.WriteLine("User has signed out");
                }

                else
                {
                    Console.WriteLine("Element not found");
                    //properties.driver.FindElement(By.XPath("//input[@id = 'btnAttendanceSignOut']")).Submit();
                }

            }
        }
        
    }
}
