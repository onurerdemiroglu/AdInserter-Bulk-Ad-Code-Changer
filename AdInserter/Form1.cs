using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using OpenQA.Selenium.Support.UI;

namespace AdInserter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;
        }
        void beklemeliclick(IWebDriver driver, int element1xpath2id3classname, string elementici)
        {
            tekrar:
            try
            {
                switch (element1xpath2id3classname)
                {
                    case 1:
                        driver.FindElement(By.XPath(elementici)).Click();
                        break;
                    case 2:
                        driver.FindElement(By.Id(elementici)).Click();
                        break;
                    case 3:
                        driver.FindElement(By.ClassName(elementici)).Click();
                        break;
                }
            }
            catch (Exception)
            {
                Thread.Sleep(300);
                goto tekrar;
            }
        }
        void beklemeliclear(IWebDriver driver, int element1xpath2id3classname, string elementici)
        {
            tekrar:
            try
            {
                switch (element1xpath2id3classname)
                {
                    case 1:
                        driver.FindElement(By.XPath(elementici)).Clear();
                        break;
                    case 2:
                        driver.FindElement(By.Id(elementici)).Clear();
                        break;
                    case 3:
                        driver.FindElement(By.ClassName(elementici)).Clear();
                        break;
                }
            }
            catch (Exception)
            {
                Thread.Sleep(300);
                goto tekrar;
            }
        }
        static void RunAsSTAThread(Action goForIt)
        {
            AutoResetEvent @event = new AutoResetEvent(false);
            Thread thread = new Thread(
                () =>
                {
                    goForIt();
                    @event.Set();
                });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            @event.WaitOne();
        }
        void beklemelisendkey(IWebDriver driver, int element1xpath2id3classname4name, string elementici, string sendkey)
        {
            tekrar:
            try
            {

                switch (element1xpath2id3classname4name)
                {
                    case 1:
                        for (int i = 0; i < sendkey.Length; i++)
                        {
                            driver.FindElement(By.XPath(elementici)).SendKeys(sendkey[i].ToString());
                        }
                        break;
                    case 2:
                        for (int i = 0; i < sendkey.Length; i++)
                        {
                            driver.FindElement(By.Id(elementici)).SendKeys(sendkey[i].ToString());
                        }
                        break;
                    case 3:
                        for (int i = 0; i < sendkey.Length; i++)
                        {
                            driver.FindElement(By.ClassName(elementici)).SendKeys(sendkey[i].ToString());
                        }
                        break;
                    case 4:
                        for (int i = 0; i < sendkey.Length; i++)
                        {
                            driver.FindElement(By.Name(elementici)).SendKeys(sendkey[i].ToString());
                        }
                        break;
                }
            }
            catch (Exception)
            {
                Thread.Sleep(300);
                goto tekrar;
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            Control.CheckForIllegalCrossThreadCalls = false;
            try
            {
                domainlistesi.Items.Clear();
                string[] lines = File.ReadAllLines(Application.StartupPath + "/Domain.txt");
                domainlistesi.Items.AddRange(lines);
            }
            catch (Exception)
            {
            }
        }

        private void domainekle_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(Application.StartupPath + @"\Domain.txt");
        }

        private void guncelle_Click(object sender, EventArgs e)
        {
            try
            {
                domainlistesi.Items.Clear();
                string[] lines = File.ReadAllLines(Application.StartupPath + "/Domain.txt");
                domainlistesi.Items.AddRange(lines);
            }
            catch (Exception)
            {
            }
        }

        private void Başlat_Click(object sender, EventArgs e)
        {
            if (wpkullanici.Text != "" && wpsifre.Text != "")
            {
                Thread thread = new Thread(new ThreadStart(adinserter));
                thread.Start();
            }
            else
            {
                MessageBox.Show("WP Kullanıcı adı şifre kontrol ediniz");
            }

        }
        void adinserter()
        {
            var chromeDriverService = ChromeDriverService.CreateDefaultService();
            chromeDriverService.HideCommandPromptWindow = true;
            var options = new ChromeOptions();
            options.AddUserProfilePreference("credentials_enable_service", false);
            options.AddUserProfilePreference("profile.password_manager_enabled", false);
            options.AddExcludedArgument("enable-automation");
            options.AddAdditionalCapability("useAutomationExtension", false);
            IWebDriver driver = new ChromeDriver(chromeDriverService, options);

            for (int i = 0; i < domainlistesi.Items.Count; i++)
            {
                string domainurl = domainlistesi.Items[i].ToString() + "/wp-admin/options-general.php?page=ad-inserter.php";
                driver.Navigate().GoToUrl(domainurl);
                tekrar:
                beklemelisendkey(driver, 2, "user_login", wpkullanici.Text);
                Thread.Sleep(500);
                string kullanicikontrol = driver.FindElement(By.Id("user_login")).GetAttribute("value"); 
                if (kullanicikontrol!=wpkullanici.Text)
                {
                    driver.FindElement(By.Id("user_login")).Clear();
                    goto tekrar;
                }
                beklemelisendkey(driver, 2, "user_pass", wpsifre.Text);
                beklemeliclick(driver, 2, "wp-submit");


                string butonid = "";
                ((IJavaScriptExecutor)driver).ExecuteScript("window.scrollTo(document.body.scrollHeight, 100)");
                if (reklam.Text != "")
                {
                    RunAsSTAThread(
                          () =>
                          {
                              Clipboard.SetText(reklam.Text);
                          }); 

                    yine:
                    try
                    {
                        driver.FindElement(By.XPath("//a[@id='ui-id-1']")).Click();
                    }
                    catch (Exception)
                    {
                        IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                        js.ExecuteScript("window.scrollBy(0,-5000)");
                        Thread.Sleep(300);
                        goto yine;
                    }

                    beklemeliclear(driver, 1, "//*[@id='editor-1']/textarea");
                    driver.FindElement(By.XPath("//*[@id='editor-1']/textarea")).SendKeys(OpenQA.Selenium.Keys.Control + "v");
                    beklemeliclick(driver, 1, ".//*[@id='block-alignment-1']/option[3]");
                    beklemeliclick(driver, 1, ".//*[@id='insertion-type-1']/option[3]");

                    SelectElement element2 = new SelectElement(driver.FindElement(By.Id("insertion-type-1")));
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2.0);
                    element2.SelectByText(reklamyeri1.Text);

                    Thread.Sleep(1000);
                    butonid = "2"; 
                }
                if (reklam2.Text != "")
                {
                    RunAsSTAThread(
                         () =>
                         {
                             Clipboard.SetText(reklam2.Text);
                         });

                    yine:
                    try
                    {
                        driver.FindElement(By.XPath("//a[@id='ui-id-2']")).Click();
                    }
                    catch (Exception)
                    {
                        IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                        js.ExecuteScript("window.scrollBy(0,-5000)");
                        Thread.Sleep(300);
                        goto yine;
                    }

                    beklemeliclear(driver, 1, "//*[@id='editor-2']/textarea");
                    driver.FindElement(By.XPath("//*[@id='editor-2']/textarea")).SendKeys(OpenQA.Selenium.Keys.Control + "v");
                    beklemeliclick(driver, 1, ".//*[@id='block-alignment-2']/option[3]");

                    SelectElement element2 = new SelectElement(driver.FindElement(By.Id("insertion-type-2")));
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2.0);
                    element2.SelectByText(reklamyeri2.Text);

                    Thread.Sleep(1000);
                    butonid = "3";
                }
                if (reklam3.Text != "")
                {
                    RunAsSTAThread(
                            () =>
                            {
                                Clipboard.SetText(reklam3.Text);
                            });

                    yine:
                    try
                    {
                        driver.FindElement(By.XPath("//a[@id='ui-id-3']")).Click();
                    }
                    catch (Exception)
                    {
                        IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                        js.ExecuteScript("window.scrollBy(0,-5000)");
                        Thread.Sleep(300);
                        goto yine;
                    }
                    beklemeliclick(driver, 1, "//a[@id='ui-id-3']");
                    beklemeliclear(driver, 1, "//*[@id='editor-3']/textarea");
                    driver.FindElement(By.XPath("//*[@id='editor-3']/textarea")).SendKeys(OpenQA.Selenium.Keys.Control + "v");
                    beklemeliclick(driver, 1, ".//*[@id='block-alignment-3']/option[3]");

                    SelectElement element2 = new SelectElement(driver.FindElement(By.Id("insertion-type-3")));
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2.0);
                    element2.SelectByText(reklamyeri3.Text);

                    Thread.Sleep(1000);
                    butonid = "4";
                }
                if (reklam4.Text != "")
                {
                    RunAsSTAThread(
                            () =>
                            {
                                Clipboard.SetText(reklam4.Text);
                            });

                    yine:
                    try
                    {
                        driver.FindElement(By.XPath("//a[@id='ui-id-4']")).Click();
                    }
                    catch (Exception)
                    {
                        IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                        js.ExecuteScript("window.scrollBy(0,-5000)");
                        Thread.Sleep(300);
                        goto yine;
                    }

                    beklemeliclear(driver, 1, "//*[@id='editor-4']/textarea");
                    driver.FindElement(By.XPath("//*[@id='editor-4']/textarea")).SendKeys(OpenQA.Selenium.Keys.Control + "v");
                    beklemeliclick(driver, 1, ".//*[@id='block-alignment-4']/option[3]");

                    SelectElement element2 = new SelectElement(driver.FindElement(By.Id("insertion-type-4")));
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2.0);
                    element2.SelectByText(reklamyeri4.Text);

                    Thread.Sleep(1000);
                    butonid = "5";
                }
                if (reklam5.Text != "")
                {
                    RunAsSTAThread(
                            () =>
                            {
                                Clipboard.SetText(reklam5.Text);
                            }); 

                    yine:
                    try
                    {
                        driver.FindElement(By.XPath("//a[@id='ui-id-5']")).Click();
                    }
                    catch (Exception)
                    {
                        IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                        js.ExecuteScript("window.scrollBy(0,-5000)");
                        Thread.Sleep(300);
                        goto yine;
                    }

                    beklemeliclear(driver, 1, "//*[@id='editor-5']/textarea");
                    driver.FindElement(By.XPath("//*[@id='editor-5']/textarea")).SendKeys(OpenQA.Selenium.Keys.Control + "v");
                    beklemeliclick(driver, 1, ".//*[@id='block-alignment-5']/option[3]");

                    SelectElement element2 = new SelectElement(driver.FindElement(By.Id("insertion-type-5")));
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2.0);
                    element2.SelectByText(reklamyeri5.Text);

                    Thread.Sleep(1000);
                    butonid = "6"; 
                }
                if (reklam6.Text != "")
                {
                    RunAsSTAThread(
                            () =>
                            {
                                Clipboard.SetText(reklam6.Text);
                            });

                    yine:
                    try
                    {
                        driver.FindElement(By.XPath("//a[@id='ui-id-6']")).Click();
                    }
                    catch (Exception)
                    {
                        IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                        js.ExecuteScript("window.scrollBy(0,-5000)");
                        Thread.Sleep(300);
                        goto yine;
                    }

                    beklemeliclear(driver, 1, "//*[@id='editor-6']/textarea");
                    driver.FindElement(By.XPath("//*[@id='editor-6']/textarea")).SendKeys(OpenQA.Selenium.Keys.Control + "v");
                    beklemeliclick(driver, 1, ".//*[@id='block-alignment-6']/option[3]");

                    SelectElement element2 = new SelectElement(driver.FindElement(By.Id("insertion-type-6")));
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2.0);
                    element2.SelectByText(reklamyeri6.Text);

                    Thread.Sleep(1000);
                    butonid = "7";
                }


                beklemeliclick(driver, 1, "(//input[@name='ai_save'])[" + butonid + "]");


                beklemeliclick(driver, 1, "(//div[@class='notice notice-success is-dismissible'])"); /// kaydedildi mi?


                Thread.Sleep(1000);
                 
                label2.Text = "Durum: " + domainlistesi.Items[i].ToString() + " adlı site yapıldı. devam ediliyor.";
            }
            driver.Quit(); 
        }
         
    }
}
