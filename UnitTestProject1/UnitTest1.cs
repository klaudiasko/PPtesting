using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Threading;

namespace UnitTestProject1
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void OpenPowerPointAndCheckItems()
        {
            var app = FlaUI.Core.Application.Launch(@"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE");
            using var automation = new UIA3Automation();
            Thread.Sleep(1000);
            var window = app.GetMainWindow(automation);
            Assert.AreEqual("PowerPoint", window.Title);
            
            var menu = window.FindFirstDescendant(x => x.ByAutomationId("NavBarMenu"))?.AsListBox();
            var menuList = new[] { "Strona g³ówna", "Nowy", "Otwórz", "Konto", "Opinia", "Opcje" };
            var i = 0;
            foreach (var item in menu.FindAllChildren())
            {
                Assert.AreEqual(menuList[i], item.AsListBoxItem().Name);
                i++;
            }

            window.Close();
            app.Close();
        }

        
        [TestMethod]
        public void OpenNewPresentationAndChangeTheme()
        {
            var app = FlaUI.Core.Application.Launch(@"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE");
            using var automation = new UIA3Automation();
            Thread.Sleep(1000);
            var window = app.GetMainWindow(automation);
           
            
            var newListItem = window.FindFirstDescendant(x => x.ByName("Pusta prezentacja"))?.AsButton();
            newListItem?.Click();
            Thread.Sleep(500);

            var projects = window.FindFirstDescendant(x => x.ByName("Projektowanie"))?.AsButton();
            projects?.Click();
            Thread.Sleep(500);

            
            var prTheme1 = window.FindFirstDescendant(x => x.ByName("Integralny Trwa ³adowanie podgl¹du..."))?.AsButton();
            prTheme1?.Click();
            Thread.Sleep(500);
            Assert.AreEqual("Integralny Trwa ³adowanie podgl¹du...", prTheme1.AsButton().Name);

            var prTheme2 = window.FindFirstDescendant(x => x.ByName("Jon (sala konferencyjna) Trwa ³adowanie podgl¹du..."))?.AsButton();
            prTheme2?.Click();
            Thread.Sleep(500);
            Assert.AreEqual("Jon (sala konferencyjna) Trwa ³adowanie podgl¹du...", prTheme2.AsButton().Name);

            var closeButton = window.FindFirstDescendant(cf => cf.ByName("Zamknij"))?.AsButton();
            closeButton?.Invoke();
            Thread.Sleep(500);

            var notSaveButton = window.FindFirstDescendant(cf => cf.ByName("Nie zapisuj"))?.AsButton();
            notSaveButton?.Invoke();

            app.Close();
            
        }
        [TestMethod]
        public void OpenNewPresentationAndChangeTransitions()
        {
            var app = FlaUI.Core.Application.Launch(@"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE");
            using var automation = new UIA3Automation();
            Thread.Sleep(1000);
            var window = app.GetMainWindow(automation);
            

            var newListItem = window.FindFirstDescendant(x => x.ByName("Pusta prezentacja"))?.AsButton();
            newListItem?.Click();
            Thread.Sleep(500);

            
            var projects = window.FindFirstDescendant(x => x.ByName("Projektowanie"))?.AsButton();
            projects?.Click();
            Thread.Sleep(500);


            var prTheme1 = window.FindFirstDescendant(x => x.ByName("Organiczny Trwa ³adowanie podgl¹du..."))?.AsButton();
            prTheme1?.Click();
            Thread.Sleep(500);

           

            var transitions = window.FindFirstDescendant(x => x.ByName("Przejœcia"))?.AsButton();
            transitions?.Click();
            Thread.Sleep(500);

            var prTransition = window.FindFirstDescendant(x => x.ByName("Kszta³t"))?.AsButton();
            prTransition?.Click();
            Thread.Sleep(500);
            Assert.AreEqual("Kszta³t", prTransition.AsButton().Name);

            var closeButton = window.FindFirstDescendant(cf => cf.ByName("Zamknij"))?.AsButton();
            closeButton?.Invoke();
            Thread.Sleep(500);

            var notSaveButton = window.FindFirstDescendant(cf => cf.ByName("Nie zapisuj"))?.AsButton();
            notSaveButton?.Invoke();

            app.Close();

        }

        [TestMethod]
        public void OpenNewPresentationAndChangeView()
        {
            var app = FlaUI.Core.Application.Launch(@"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE");
            using var automation = new UIA3Automation();
            Thread.Sleep(1000);
            var window = app.GetMainWindow(automation);
            

            var newListItem = window.FindFirstDescendant(x => x.ByName("Pusta prezentacja"))?.AsButton();
            newListItem?.Invoke();
            Thread.Sleep(500);

             
            var slajd = window.FindFirstDescendant(x => x.ByName("Nowy slajd"))?.AsButton();
            slajd?.Invoke();
            Thread.Sleep(500);

            var view = window.FindFirstDescendant(x => x.ByName("Widok"))?.AsButton();
            view?.Click();
            Thread.Sleep(500);

            var outlineView = window.FindFirstDescendant(x => x.ByName("Widok konspektu"))?.AsButton();
            outlineView?.Click();
            Thread.Sleep(500);
            Assert.AreEqual("Widok konspektu", outlineView.AsButton().Name);



            var closeButton = window.FindFirstDescendant(cf => cf.ByName("Zamknij"))?.AsButton();
            closeButton?.Click();
            Thread.Sleep(500);

            var notSaveButton = window.FindFirstDescendant(cf => cf.ByName("Nie zapisuj"))?.AsButton();
            notSaveButton?.Invoke();

            app.Close();

        }


    }
}