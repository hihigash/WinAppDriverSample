using System;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Remote;

namespace WinAppDriverSample
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            var screenshotFolder = @"C:\works\screenshot";

            var appCapabilities = new DesiredCapabilities();

            // PC 設定アプリを起動
            var identifier = "Windows.ImmersiveControlPanel_cw5n1h2txyewy!microsoft.windows.immersivecontrolpanel";
            appCapabilities.SetCapability("app", identifier);
            var session = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4732"), appCapabilities);
            Assert.IsNotNull(session);

            session.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(5));

            session.GetScreenshot().SaveAsFile(Path.Combine(screenshotFolder, "Settings.jpg"), ImageFormat.Jpeg);

            var items = session.FindElementByAccessibilityId("PageGroupsListView")
                .FindElementsByClassName("GridViewItem").Select(x => x.Text).ToList();
            foreach (var item in items)
            {
                Console.WriteLine(item);
                // NOTE: 画面遷移ごとに DOM ツリーが更新されるので、セッションから再取得する
                session.FindElementByAccessibilityId("PageGroupsListView").FindElementByName(item).Click();
                var pageListView = session.FindElementsByAccessibilityId("PagesListView").FirstOrDefault();
                var submenuItems = pageListView?.FindElementsByClassName("ListViewItem")?.Select(x => x.Text);
                if (submenuItems != null)
                {
                    foreach (var submenuItemText in submenuItems)
                    {
                        Console.WriteLine($"\t{submenuItemText}");
                        session.FindElementByAccessibilityId("PagesListView").FindElementByName(submenuItemText).Click();
                        Thread.Sleep(TimeSpan.FromSeconds(3));

                        var screenshot = session.GetScreenshot();
                        screenshot.SaveAsFile(Path.Combine(screenshotFolder, $"{item}-{submenuItemText}.jpg"), ImageFormat.Jpeg);

                    }
                }

                session.FindElementByAccessibilityId("Back").Click(); // [戻る] をクリック
            }

            // アプリを終了する
            session.Quit();
        }
    }
}
