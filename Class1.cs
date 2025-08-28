using System;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Drawing;                 // Point, Rectangle
using System.Runtime.InteropServices; // COMException
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Input;
using FlaUI.Core.Tools;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;
using NUnit.Framework;

namespace MyFlaUITests2
{
    [TestFixture]
    [Apartment(System.Threading.ApartmentState.STA)]
    [NonParallelizable]
    public class MyAppTest3
    {
        private Application app;
        private UIA3Automation automation;

        [SetUp]
        public void Setup()
        {
            var psi = new ProcessStartInfo
            {
                FileName = @"C:\Program Files\Autodesk\Revit 2026\Revit.exe",
            };

            // 1) Launch / Attach
            Application.Launch(psi);
            Thread.Sleep(120000); // có thể tăng nếu máy mở Revit lâu
            app = Application.AttachOrLaunch(psi);

            automation = new UIA3Automation();
            var main = app.GetMainWindow(automation, TimeSpan.FromSeconds(120));
            if (main == null)
            {
                Assert.Fail("Revit failed to start or exited immediately.");
                return;
            }

            main.SetForeground();
            Wait.UntilInputIsProcessed();

            // 2) Chỉ đóng Trial bằng IPM Loader
            ClickTrialCloseFromIpmLoader(main, TimeSpan.FromSeconds(15));
        }

        // =============== TEST: Open Recent -> verify plugin loaded ===============
        [Test]
        public void Test_OpenRecent_Then_Verify_ElectricalLoadPlugin_Loaded()
        {
            const string recentTitle = "Project1";
            const string pluginButtonName = "ElectricalLoadPlugin";

            var mainWindow = app.GetMainWindow(automation, TimeSpan.FromSeconds(60));
            Assert.That(mainWindow, Is.Not.Null, "Không tìm thấy Autodesk Revit main window");

            // ---- B1: Mở file Project1 từ Recent ----
            var opened = OpenRecentProjectByDataItem(mainWindow, recentTitle, TimeSpan.FromSeconds(30));
            Assert.That(opened, Is.True, $"Không thể mở Project '{recentTitle}' từ Recent.");

            // ---- B2: Kiểm tra plugin có mặt ----
            var pluginBtn = FindRibbonButton(mainWindow, pluginButtonName, TimeSpan.FromSeconds(10));
            Assert.That(pluginBtn, Is.Not.Null, "Không tìm thấy plugin ElectricalLoadPlugin");
            Assert.That(pluginBtn.IsEnabled, Is.True, "Plugin ElectricalLoadPlugin đang bị disable.");

            // ---- B3: Click plugin để mở group con ----
            SafeSelectOrInvoke(pluginBtn);
            Wait.UntilInputIsProcessed();
            Thread.Sleep(500);

            // ---- B4: Kiểm tra 2 nút "Execute" và "Insert Equipment" xuất hiện ----
            var executeBtn = FindRibbonButton(mainWindow, "Execute", TimeSpan.FromSeconds(5));
            var insertBtn = FindRibbonButton(mainWindow, "Insert Equipment", TimeSpan.FromSeconds(5));

            Assert.That(executeBtn, Is.Not.Null, "Không tìm thấy nút Execute trong plugin.");
            Assert.That(insertBtn, Is.Not.Null, "Không tìm thấy nút Insert Equipment trong plugin.");

            // ---- B5: Click Execute và kiểm tra dialog ----
            SafeSelectOrInvoke(executeBtn);
            Wait.UntilInputIsProcessed();

            var dialog = WaitForDialogByTitle("ElectricalLoadPlugin", TimeSpan.FromSeconds(20));
            Thread.Sleep(5000);
            Assert.That(dialog, Is.Not.Null, "Dialog ElectricalLoadPlugin không xuất hiện sau khi click Execute.");
        }

        // ---------------- OPEN RECENT bằng DataItem (50029) ----------------
        private bool OpenRecentProjectByDataItem(Window mainWindow, string projectName, TimeSpan timeout)
        {
            var cf = automation.ConditionFactory;
            var deadline = DateTime.UtcNow + timeout;

            AutomationElement item = null;

            // Nếu có ô search (giống ảnh), gõ tên để lọc
            try
            {
                var searchBox = mainWindow.FindFirstDescendant(
                    cf.ByControlType(FlaUI.Core.Definitions.ControlType.Edit)
                      .And(cf.ByName("Project1").Not())); // tránh trúng textbox đã hiển thị text như ảnh
                // Nếu không chắc, có thể tìm theo "Search for recent files"
                if (searchBox == null)
                {
                    searchBox = mainWindow.FindFirstDescendant(
                        cf.ByControlType(FlaUI.Core.Definitions.ControlType.Edit)
                          .And(cf.ByName("Search for recent files")));
                }
                if (searchBox != null)
                {
                    var tb = searchBox.AsTextBox();
                    tb.Text = projectName;
                    Keyboard.Press(VirtualKeyShort.RETURN);
                    Wait.UntilInputIsProcessed();
                    Thread.Sleep(600);
                }
            }
            catch { }

            while (DateTime.UtcNow < deadline && item == null)
            {
                try
                {
                    // Ưu tiên đúng ControlType = DataItem (50029)
                    item = mainWindow.FindFirstDescendant(
                        cf.ByControlType(FlaUI.Core.Definitions.ControlType.DataItem)
                          .And(cf.ByName(projectName)));

                    // Fallback 1: tìm Text con rồi lấy parent row (một số theme)
                    if (item == null)
                    {
                        var text = mainWindow.FindFirstDescendant(
                            cf.ByControlType(FlaUI.Core.Definitions.ControlType.Text)
                              .And(cf.ByName(projectName)));
                        if (text != null) item = text.Parent;
                    }

                    // Fallback 2: tìm ListItem/Pane nếu UI khác
                    if (item == null)
                    {
                        item = mainWindow.FindFirstDescendant(
                            cf.ByName(projectName)
                              .And(cf.ByControlType(FlaUI.Core.Definitions.ControlType.ListItem)
                                   .Or(cf.ByControlType(FlaUI.Core.Definitions.ControlType.Pane))));
                    }
                }
                catch (COMException) { }
                catch { }

                if (item == null) Thread.Sleep(150);
            }

            if (item == null) return false;

            // Scroll vào view nếu cần
            try { item.Patterns.ScrollItem.PatternOrDefault?.ScrollIntoView(); } catch { }

            // DoubleClick vào giữa DataItem để mở
            try
            {
                var r = item.BoundingRectangle;
                var p = new Point((int)(r.X + r.Width / 2), (int)(r.Y + r.Height / 2));
                mainWindow.SetForeground();
                Wait.UntilInputIsProcessed();
                Mouse.MoveTo(p);
                Thread.Sleep(80);
                Mouse.DoubleClick();
            }
            catch { }

            // Nếu vẫn chưa mở, thử Invoke trên cell con (nếu có)
            if (!WaitUntilProjectOpen(projectName, TimeSpan.FromSeconds(6)))
            {
                try
                {
                    var firstCell = item.FindFirstDescendant(cf.ByControlType(FlaUI.Core.Definitions.ControlType.Text))
                               ?? item.FindFirstDescendant(cf.ByControlType(FlaUI.Core.Definitions.ControlType.Image))
                               ?? item;
                    if (firstCell.Patterns.Invoke.IsSupported)
                    {
                        firstCell.Patterns.Invoke.Pattern.Invoke();
                    }
                    else
                    {
                        var rr = firstCell.BoundingRectangle;
                        var q = new Point((int)(rr.X + Math.Min(40, rr.Width / 3)), (int)(rr.Y + rr.Height / 2));
                        Mouse.MoveTo(q);
                        Thread.Sleep(60);
                        Mouse.DoubleClick();
                    }
                }
                catch { }
            }

            return WaitUntilProjectOpen(projectName, TimeSpan.FromSeconds(8));
        }

        private bool WaitUntilProjectOpen(string recentTitle, TimeSpan wait)
        {
            var stop = DateTime.UtcNow + wait;
            while (DateTime.UtcNow < stop)
            {
                try
                {
                    var main = app.GetMainWindow(automation, TimeSpan.FromSeconds(2));
                    if (main != null && IsProjectOpen(main, recentTitle)) return true;
                }
                catch { }
                Thread.Sleep(300);
            }
            return false;
        }

        private bool IsProjectOpen(Window mainWindow, string recentTitle)
        {
            // 1) Title chứa tên project
            try
            {
                var title = mainWindow.Title;
                if (!string.IsNullOrEmpty(title) &&
                    title.IndexOf(recentTitle, StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            catch { }

            // 2) TabItem chứa tên project
            try
            {
                var cf = automation.ConditionFactory;
                var tabs = mainWindow.FindAllDescendants(
                    cf.ByControlType(FlaUI.Core.Definitions.ControlType.TabItem));
                if (tabs != null && tabs.Length > 0)
                {
                    foreach (var t in tabs)
                    {
                        var name = SafeGetName(t);
                        if (!string.IsNullOrEmpty(name) &&
                            name.IndexOf(recentTitle, StringComparison.OrdinalIgnoreCase) >= 0)
                            return true;
                    }
                }
            }
            catch { }

            // 3) Text control chứa tên file
            try
            {
                var cf = automation.ConditionFactory;
                var texts = mainWindow.FindAllDescendants(
                    cf.ByControlType(FlaUI.Core.Definitions.ControlType.Text));
                if (texts != null)
                {
                    foreach (var tx in texts)
                    {
                        var n = SafeGetName(tx);
                        if (!string.IsNullOrEmpty(n) &&
                            n.IndexOf(recentTitle, StringComparison.OrdinalIgnoreCase) >= 0)
                            return true;
                    }
                }
            }
            catch { }

            return false;
        }
        private Window WaitForDialogByTitle(string title, TimeSpan timeout)
        {
            var stop = DateTime.UtcNow + timeout;
            while (DateTime.UtcNow < stop)
            {
                try
                {
                    var desktop = automation.GetDesktop();
                    var dialog = desktop.FindFirstDescendant(
                        automation.ConditionFactory.ByControlType(FlaUI.Core.Definitions.ControlType.Window)
                        .And(automation.ConditionFactory.ByName(title)));
                    if (dialog != null) return dialog.AsWindow();
                }
                catch { }

                Thread.Sleep(300);
            }
            return null;
        }
        // ---------------- Trial Popup (IPM Loader) ----------------
        private bool ClickTrialCloseFromIpmLoader(Window mainWindow, TimeSpan timeout)
        {
            var cf = automation.ConditionFactory;
            var start = DateTime.UtcNow;

            while (DateTime.UtcNow - start < timeout)
            {
                mainWindow.SetForeground();
                Wait.UntilInputIsProcessed();
                Thread.Sleep(150);

                var ipmDoc = mainWindow.FindFirstDescendant(
                    cf.ByControlType(FlaUI.Core.Definitions.ControlType.Document)
                      .And(cf.ByName("IPM Loader")));

                if (ipmDoc == null)
                {
                    Thread.Sleep(300);
                    continue;
                }

                var links = ipmDoc.FindAllDescendants(
                    cf.ByControlType(FlaUI.Core.Definitions.ControlType.Hyperlink));
                if (links == null || links.Length == 0)
                {
                    Thread.Sleep(200);
                    continue;
                }

                var docRect = ipmDoc.BoundingRectangle;
                var topBand = docRect.Top + 80;
                var rightBand = docRect.Right - 120;

                AutomationElement candidate = null;
                foreach (var el in links)
                {
                    var r = el.BoundingRectangle;
                    var hasInvoke = el.Patterns.Invoke.IsSupported ||
                                    el.Patterns.LegacyIAccessible.IsSupported;

                    if (r.Top <= topBand && r.Right >= rightBand &&
                        r.Width <= 60 && r.Height <= 60 &&
                        !el.IsOffscreen && hasInvoke)
                    {
                        if (candidate == null) candidate = el;
                        else
                        {
                            var cr = candidate.BoundingRectangle;
                            var curArea = r.Width * r.Height;
                            var bestArea = cr.Width * cr.Height;
                            if (curArea < bestArea || (Math.Abs(curArea - bestArea) < 0.1 && r.Right > cr.Right))
                                candidate = el;
                        }
                    }
                }

                if (candidate != null)
                {
                    try
                    {
                        if (candidate.Patterns.Invoke.IsSupported)
                            candidate.Patterns.Invoke.Pattern.Invoke();
                        else if (candidate.Patterns.LegacyIAccessible.IsSupported)
                            candidate.Patterns.LegacyIAccessible.Pattern.DoDefaultAction();
                        else
                        {
                            var r = candidate.BoundingRectangle;
                            var p = new Point((int)(r.X + r.Width / 2), (int)(r.Y + r.Height / 2));
                            Mouse.MoveTo(p);
                            Thread.Sleep(60);
                            Mouse.Click();
                        }

                        Wait.UntilInputIsProcessed();
                        Thread.Sleep(250);

                        var stillThere = mainWindow.FindFirstDescendant(
                            cf.ByControlType(FlaUI.Core.Definitions.ControlType.Document)
                              .And(cf.ByName("IPM Loader")));
                        if (stillThere == null) return true;
                    }
                    catch { }
                }
                else
                {
                    Thread.Sleep(200);
                }
            }
            return false;
        }

        // ---------------- Helpers an toàn chống COMException ----------------
        private static string SafeGetName(AutomationElement el)
        {
            if (el == null) return null;
            try
            {
                string n;
                if (el.Properties.Name.TryGetValue(out n)) return n;
            }
            catch (COMException) { }
            catch { }
            try { return el.Name; } catch { return null; }
        }

        private static bool SafeIsOffscreen(AutomationElement el)
        {
            try { return el.IsOffscreen; } catch { return false; }
        }

        private static void SafeSelectOrInvoke(AutomationElement el)
        {
            if (el == null) return;
            try
            {
                var sel = el.Patterns.SelectionItem.PatternOrDefault;
                if (sel != null) { sel.Select(); return; }
            }
            catch { }
            try
            {
                if (el.Patterns.Invoke.IsSupported) { el.Patterns.Invoke.Pattern.Invoke(); return; }
            }
            catch { }
            try { el.Click(); } catch { }
        }

        private AutomationElement FindRibbonButton(Window mainWindow, string buttonName, TimeSpan timeout)
        {
            var cf = automation.ConditionFactory;
            var start = DateTime.UtcNow;

            while (DateTime.UtcNow - start < timeout)
            {
                try
                {
                    mainWindow.SetForeground();
                    Wait.UntilInputIsProcessed();

                    var btn = mainWindow.FindFirstDescendant(
                        cf.ByControlType(FlaUI.Core.Definitions.ControlType.Button)
                          .And(cf.ByName(buttonName)));

                    if (btn != null && !SafeIsOffscreen(btn)) return btn;
                }
                catch (COMException) { }
                catch { }

                Thread.Sleep(200);
            }
            return null;
        }

    
        [TearDown]
        public void Teardown()
        {
            try
            {
                if (app != null && !app.HasExited)
                {
                    app.Close();
                    Thread.Sleep(500);
                }
            }
            catch { }

            foreach (var p in Process.GetProcessesByName("Revit"))
            {
                try
                {
                    if (!p.HasExited)
                    {
                        p.CloseMainWindow();
                        p.WaitForExit(1000);
                    }
                }
                catch { }
            }

            automation?.Dispose();
        }
    }
}
