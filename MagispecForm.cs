//-----------------------------------------------------------------------
// <copyright file="MagispecForm.cs" company="Andy Young">
//     Copyright (c) Andy Young. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------
namespace Magispec
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Drawing;
    using System.Linq;
    using System.Windows.Forms;
    using Microsoft.Win32;

    /// <summary>
    /// Magispec form
    /// </summary>
    public partial class MagispecForm : Form
    {
        /// <summary>
        /// Map between each trait and its background color
        /// </summary>
        private static readonly Dictionary<Trait, Color> TraitColors = new Dictionary<Trait, Color>()
        {
            { Trait.Agressive, Color.FromArgb(157, 39, 42) },
            { Trait.Artisan, Color.FromArgb(150, 62, 40) },
            { Trait.BigEater, Color.FromArgb(127, 67, 39) },
            { Trait.Defensive, Color.FromArgb(49, 106, 158) },
            { Trait.Gatherer, Color.FromArgb(86, 121, 67) },
            { Trait.Healthy, Color.FromArgb(234, 37, 118) },
            { Trait.Intellect, Color.FromArgb(136, 54, 170) },
            { Trait.Lockmaster, Color.FromArgb(118, 119, 117) },
            { Trait.Miner, Color.FromArgb(90, 88, 92) },
            { Trait.PotionBrewer, Color.FromArgb(108, 31, 153) },
            { Trait.Swift, Color.FromArgb(35, 152, 73) },
            { Trait.Woodcutter, Color.FromArgb(126, 82, 63) }
        };

        /// <summary>
        /// Positions for the "Stats" button for each window resolution
        /// </summary>
        private static readonly Dictionary<string, Rectangle> StatsButtonPosition = new Dictionary<string, Rectangle>()
        {
            { "640x480", new Rectangle(434, 127, 104, 29) },
            { "800x600", new Rectangle(543, 159, 129, 36) },
            { "1024x768", new Rectangle(695, 204, 165, 46) },
            { "1152x864", new Rectangle(781, 230, 187, 51) },
            { "1280x720", new Rectangle(847, 156, 187, 52) },
            { "1280x960", new Rectangle(868, 255, 207, 57) },
            { "1280x1024", new Rectangle(872, 284, 211, 57) },
            { "1366x768", new Rectangle(904, 167, 199, 55) },
            { "1400x1050", new Rectangle(949, 279, 227, 63) },
            { "1600x900", new Rectangle(1058, 196, 235, 64) },
            { "1600x1024", new Rectangle(1071, 245, 246, 68) },
            { "1920x1080", new Rectangle(1270, 235, 281, 77) }
        };

        /// <summary>
        /// Positions for the first trait icon for each window resolution
        /// </summary>
        private static readonly Dictionary<string, Rectangle> Trait1IconPosition = new Dictionary<string, Rectangle>()
        {
            { "640x480", new Rectangle(436, 263, 50, 50) },
            { "800x600", new Rectangle(545, 329, 62, 62) },
            { "1024x768", new Rectangle(698, 421, 80, 80) },
            { "1152x864", new Rectangle(785, 474, 90, 90) },
            { "1280x720", new Rectangle(850, 402, 90, 90) },
            { "1280x960", new Rectangle(872, 527, 99, 99) },
            { "1280x1024", new Rectangle(876, 560, 101, 101) },
            { "1366x768", new Rectangle(907, 429, 96, 96) },
            { "1400x1050", new Rectangle(954, 576, 109, 109) },
            { "1600x900", new Rectangle(1063, 503, 113, 113) },
            { "1600x1024", new Rectangle(1076, 568, 118, 118) },
            { "1920x1080", new Rectangle(1275, 604, 135, 135) }
        };

        /// <summary>
        /// Positions for the second trait icon for each window resolution
        /// </summary>
        private static readonly Dictionary<string, Rectangle> Trait2IconPosition = BuildTrait2IconPositions();

        /// <summary>
        /// The Magicite registry key
        /// </summary>
        private RegistryKey regKey;

        /// <summary>
        /// The list of the value names of the Magicite registry key
        /// </summary>
        private string[] regkeyValueNames;

        /// <summary>
        /// The registry key value which knows if the game is full-screen
        /// </summary>
        private string isFullScreenRegKeyValue;

        /// <summary>
        /// The registry key value which knows the window resolution height
        /// </summary>
        private string resolutionHeightValue;

        /// <summary>
        /// The registry key value which knows the window resolution width
        /// </summary>
        private string resolutionWidthValue;

        /// <summary>
        /// List of the checkboxes
        /// </summary>
        private CheckBox[] checkboxes;

        /// <summary>
        /// Initializes a new instance of the <see cref="MagispecForm" /> class.
        /// </summary>
        public MagispecForm()
        {
            this.InitializeComponent();
        }

        /// <summary>
        /// Gets the Magicite registry key
        /// </summary>
        private RegistryKey RegKey
        {
            get
            {
                if (this.regKey == null)
                {
                    this.regKey = Registry.CurrentUser.OpenSubKey(@"Software\SmashGames\Magicite");
                    if (this.regKey == null)
                    {
                        throw new ApplicationException(@"Cannot determine Magicite registry key location. Looked in HKCU\Software\SmashGames\Magicite");
                    }
                }

                return this.regKey;
            }
        }

        /// <summary>
        /// Gets the list of the value names of the Magicite registry key
        /// </summary>
        private string[] RegKeyValueNames
        {
            get
            {
                return this.regkeyValueNames ?? (this.regkeyValueNames = this.RegKey.GetValueNames());
            }
        }

        /// <summary>
        /// Gets the registry key value which knows if the game is full-screen
        /// </summary>
        private string IsFullScreenRegKeyValue
        {
            get
            {
                if (this.isFullScreenRegKeyValue == null)
                {
                    var query = this.RegKeyValueNames.Where(k => k.StartsWith("Screenmanager Is Fullscreen mode"));
                    if (query.Any())
                    {
                        this.isFullScreenRegKeyValue = query.First();
                    }
                }

                return this.isFullScreenRegKeyValue;
            }
        }

        /// <summary>
        /// Gets a value indicating whether the window is full-screen
        /// </summary>
        private bool IsFullScreen
        {
            get
            {
                bool isFullScreen = false;

                if (!string.IsNullOrEmpty(this.IsFullScreenRegKeyValue))
                {
                    object o = this.RegKey.GetValue(this.IsFullScreenRegKeyValue);
                    isFullScreen = o != null && int.Equals(o, 1);
                }

                return isFullScreen;
            }
        }

        /// <summary>
        /// Gets the registry key value which knows the window resolution height
        /// </summary>
        private string ResolutionHeightValue
        {
            get
            {
                if (this.resolutionHeightValue == null)
                {
                    var query = this.RegKeyValueNames.Where(k => k.StartsWith("Screenmanager Resolution Height"));
                    if (query.Any())
                    {
                        this.resolutionHeightValue = query.First();
                    }
                }

                return this.resolutionHeightValue;
            }
        }

        /// <summary>
        /// Gets the window resolution height
        /// </summary>
        private int ResolutionHeight
        {
            get
            {
                int height = -1;
                if (!string.IsNullOrEmpty(this.ResolutionHeightValue))
                {
                    object o = this.RegKey.GetValue(this.ResolutionHeightValue);
                    if (o is int)
                    {
                        height = (int)o;
                    }
                }

                return height;
            }
        }

        /// <summary>
        /// Gets the registry key value which knows the window resolution width
        /// </summary>
        private string ResolutionWidthValue
        {
            get
            {
                if (this.resolutionWidthValue == null)
                {
                    var query = this.RegKeyValueNames.Where(k => k.StartsWith("Screenmanager Resolution Width"));
                    if (query.Any())
                    {
                        this.resolutionWidthValue = query.First();
                    }
                }

                return this.resolutionWidthValue;
            }
        }

        /// <summary>
        /// Gets the window resolution width
        /// </summary>
        private int ResolutionWidth
        {
            get
            {
                int width = -1;
                if (!string.IsNullOrEmpty(this.ResolutionWidthValue))
                {
                    object o = this.RegKey.GetValue(this.ResolutionWidthValue);
                    if (o is int)
                    {
                        width = (int)o;
                    }
                }

                return width;
            }
        }

        /// <summary>
        /// Gets the window resolution string: e.g. "1024x768"
        /// </summary>
        private string ResolutionString
        {
            get
            {
                return string.Format("{0}x{1}", this.ResolutionWidth, this.ResolutionHeight);
            }
        }

        /// <summary>
        /// Builds the positions for the second trait icon for each window resolution based on those of the first icon
        /// </summary>
        /// <returns>See summary</returns>
        private static Dictionary<string, Rectangle> BuildTrait2IconPositions()
        {
            var positions = new Dictionary<string, Rectangle>();
            foreach (var kvp in Trait1IconPosition)
            {
                var trait1 = Trait1IconPosition[kvp.Key];
                positions[kvp.Key] = new Rectangle(trait1.X + trait1.Width, trait1.Y, trait1.Width, trait1.Height);
            }

            return positions;
        }

        /// <summary>
        /// Gets the Magicite process
        /// </summary>
        /// <returns>Magicite process</returns>
        private Process GetMagiciteProcess()
        {
            while (true)
            {
                var processes = Process.GetProcessesByName("magicite");
                if (processes.Length == 0)
                {
                    var result = MessageBox.Show("Please open Magicite to the trait selection screen.", "Error", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                    if (result == DialogResult.Cancel)
                    {
                        return null;
                    }
                }
                else
                {
                    return processes[0];
                }
            }
        }

        /// <summary>
        /// Updates the details of the Magicite window and process
        /// </summary>
        /// <param name="p">Magicite process</param>
        private void UpdateDetails(Process p)
        {
            var rect = this.GetWindowDimensions(p);
            listBoxDetails.Items.Clear();

            listBoxDetails.Items.Add(string.Format("Fullscreen: {0}", this.IsFullScreen));
            listBoxDetails.Items.Add(string.Format("Resolution: {0}x{1}", this.ResolutionWidth, this.ResolutionHeight));
            listBoxDetails.Items.Add(string.Format("Dimensions: {0}x{1}", rect.Width - rect.X, rect.Height - rect.Y));
            listBoxDetails.Items.Add(string.Format("Position: [{0},{1}] - [{2},{3}]", rect.Left, rect.Top, rect.Right, rect.Bottom));
            listBoxDetails.Items.Add(string.Format("Handle: {0}", p.MainWindowHandle));
        }

        /// <summary>
        /// Gets the dimensions of the window
        /// </summary>
        /// <param name="p">Magicite process</param>
        /// <returns>Dimensions of the window</returns>
        private Rectangle GetWindowDimensions(Process p)
        {
            Rectangle rect = new Rectangle();
            Win32.GetWindowRect(p.MainWindowHandle, ref rect);
            return rect;
        }

        /// <summary>
        /// Magispec form load event
        /// </summary>
        /// <param name="sender">What raised the event</param>
        /// <param name="e">Event arguments</param>
        private void MagispecForm_Load(object sender, EventArgs e)
        {
            var p = this.GetMagiciteProcess();
            if (p != null)
            {
                Timer t = new Timer() { Interval = 100 };
                t.Tick += (tickSender, tickEventArgs) =>
                {
                    p = GetMagiciteProcess();
                    if (p != null)
                    {
                        UpdateDetails(p);
                    }
                };

                t.Start();
            }

            checkBoxAggressive.Tag = Trait.Agressive;
            checkBoxArtisan.Tag = Trait.Artisan;
            checkBoxBigEater.Tag = Trait.BigEater;
            checkBoxDefensive.Tag = Trait.Defensive;
            checkBoxGatherer.Tag = Trait.Gatherer;
            checkBoxHealthy.Tag = Trait.Healthy;
            checkBoxIntelligent.Tag = Trait.Intellect;
            checkBoxLockmaster.Tag = Trait.Lockmaster;
            checkBoxMiner.Tag = Trait.Miner;
            checkBoxPotionBrewer.Tag = Trait.PotionBrewer;
            checkBoxSwift.Tag = Trait.Swift;
            checkBoxWoodcutter.Tag = Trait.Woodcutter;

            this.checkboxes = new[]
            {
                checkBoxAggressive,
                checkBoxArtisan,
                checkBoxBigEater,
                checkBoxDefensive,
                checkBoxGatherer,
                checkBoxHealthy,
                checkBoxIntelligent,
                checkBoxLockmaster,
                checkBoxMiner,
                checkBoxPotionBrewer,
                checkBoxSwift,
                checkBoxWoodcutter
            };

            foreach (var checkbox in this.checkboxes)
            {
                checkbox.ForeColor = TraitColors[(Trait)checkbox.Tag];
                checkbox.CheckedChanged += (checkedSender, checkedEventArgs) =>
                {
                    int numChecked = checkboxes.Count(c => c.Checked);
                    buttonSpec.Enabled = numChecked > 0;
                    foreach (var c in checkboxes)
                    {
                        c.Enabled = c.Checked || numChecked < 2;
                    }
                };
            }
        }

        /// <summary>
        /// Click event for the spec button
        /// </summary>
        /// <param name="sender">What raised the event</param>
        /// <param name="e">Event arguments</param>
        private void ButtonSpec_Click(object sender, EventArgs e)
        {
            var p = this.GetMagiciteProcess();
            if (p == null)
            {
                return;
            }

            Win32.ShowWindow(p.MainWindowHandle, ShowWindowCommands.Restore);
            Win32.SetForegroundWindow(p.MainWindowHandle);
            this.ClickStatsButton(p);
        }

        /// <summary>
        /// Clicks the "Stats" button until the desired traits are rolled
        /// </summary>
        /// <param name="p">Magicite process</param>
        private void ClickStatsButton(Process p)
        {
            this.TopMost = false;
            Trait desiredTrait1 = Trait.Unknown;
            Trait desiredTrait2 = Trait.Unknown;
            this.GetDesiredTraits(ref desiredTrait1, ref desiredTrait2);

            Rectangle windowRect = this.GetWindowRect(p.MainWindowHandle);
            Rectangle statsButtonRect = this.GetStatsButtonPosition(windowRect);
            this.MoveMouse(windowRect, statsButtonRect);

            Func<bool> stopCondition = new Func<bool>(() =>
            {
                if (desiredTrait1 == Trait.Unknown && desiredTrait2 == Trait.Unknown)
                {
                    return true;
                }

                Trait trait1 = Trait.Unknown;
                Trait trait2 = Trait.Unknown;
                Color background1 = GetIconBackgroundColor(GetTrait1IconPosition(windowRect), windowRect);
                Color background2 = GetIconBackgroundColor(GetTrait2IconPosition(windowRect), windowRect);
                foreach (var kvp in TraitColors)
                {
                    double similarity = 100;
                    Color traitBackround = kvp.Value;

                    if (trait1 == Trait.Unknown)
                    {
                        similarity = background1.CompareTo(traitBackround);
                        if (similarity < 11)
                        {
                            trait1 = kvp.Key;
                        }
                    }

                    if (trait2 == Trait.Unknown)
                    {
                        similarity = background2.CompareTo(traitBackround);
                        if (similarity < 11)
                        {
                            trait2 = kvp.Key;
                        }
                    }

                    if (((desiredTrait1 == Trait.Unknown || trait1 == desiredTrait1) && (desiredTrait2 == Trait.Unknown || trait2 == desiredTrait2)) ||
                        ((desiredTrait2 == Trait.Unknown || trait1 == desiredTrait2) && (desiredTrait1 == Trait.Unknown || trait2 == desiredTrait1)))
                    {
                        break;
                    }
                }

                return ((desiredTrait1 == Trait.Unknown || trait1 == desiredTrait1) && (desiredTrait2 == Trait.Unknown || trait2 == desiredTrait2)) ||
                       ((desiredTrait2 == Trait.Unknown || trait1 == desiredTrait2) && (desiredTrait1 == Trait.Unknown || trait2 == desiredTrait1));
            });

            int x = windowRect.Left + statsButtonRect.Left + ((statsButtonRect.Right - statsButtonRect.Left) / 2);
            int y = windowRect.Top + statsButtonRect.Top + ((statsButtonRect.Bottom - statsButtonRect.Top) / 2);
            this.ClickMouse(x, y, stopCondition: stopCondition);
        }

        /// <summary>
        /// Gets the Magicite window
        /// </summary>
        /// <param name="handle">Magicite handle</param>
        /// <param name="takeDebugScreenshot">True to take a screenshot of the window</param>
        /// <returns>Position of window</returns>
        private Rectangle GetWindowRect(IntPtr handle, bool takeDebugScreenshot = false)
        {
            Rectangle windowRect = Rectangle.Empty;
            Rectangle clientRect = Rectangle.Empty;
            Win32.GetWindowRect(handle, ref windowRect);
            Win32.GetClientRect(handle, ref clientRect);

            int borderWidth = (windowRect.Width - windowRect.Left - clientRect.Width) / 2;
            int titlebarHeight = windowRect.Height - windowRect.Top - clientRect.Height - borderWidth;

            // Shift the rectangle to include only the client piece
            windowRect = new Rectangle(
                x: windowRect.Left + borderWidth,
                y: windowRect.Top + titlebarHeight,
                width: windowRect.Width - borderWidth,
                height: windowRect.Height - borderWidth);

            if (takeDebugScreenshot)
            {
                Bitmap bitmap = new Bitmap(windowRect.Width - windowRect.Left, windowRect.Height - windowRect.Top);
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    g.CopyFromScreen(windowRect.Left, windowRect.Top, 0, 0, windowRect.Size);
                }

                bitmap.Save(this.ResolutionString + ".bmp", System.Drawing.Imaging.ImageFormat.Bmp);
            }

            return windowRect;
        }

        /// <summary>
        /// Gets the position of the "Stats" button
        /// </summary>
        /// <param name="windowRect">Magicite window</param>
        /// <returns>Position of the "Stats" button</returns>
        private Rectangle GetStatsButtonPosition(Rectangle windowRect)
        {
            Rectangle statsButtonRect = StatsButtonPosition[this.ResolutionString];
            if (this.IsFullScreen)
            {
                statsButtonRect = this.FullScreenTranslate(windowRect, statsButtonRect);
            }

            return statsButtonRect;
        }

        /// <summary>
        /// Gets the position of the first trait
        /// </summary>
        /// <param name="windowRect">Magicite window</param>
        /// <returns>Position of the first trait</returns>
        private Rectangle GetTrait1IconPosition(Rectangle windowRect)
        {
            Rectangle trait1IconRect = Trait1IconPosition[this.ResolutionString];
            if (this.IsFullScreen)
            {
                trait1IconRect = this.FullScreenTranslate(windowRect, trait1IconRect);
            }

            return trait1IconRect;
        }

        /// <summary>
        /// Gets the position of the second trait
        /// </summary>
        /// <param name="windowRect">Magicite window</param>
        /// <returns>Position of the second trait</returns>
        private Rectangle GetTrait2IconPosition(Rectangle windowRect)
        {
            Rectangle trait2IconRect = Trait2IconPosition[this.ResolutionString];
            if (this.IsFullScreen)
            {
                trait2IconRect = this.FullScreenTranslate(windowRect, trait2IconRect);
            }

            return trait2IconRect;
        }

        /// <summary>
        /// Translates the position of the element if in full-screen mode
        /// </summary>
        /// <param name="windowRect">Magicite window</param>
        /// <param name="rect">Element to translate</param>
        /// <returns>Translated position of element</returns>
        private Rectangle FullScreenTranslate(Rectangle windowRect, Rectangle rect)
        {
            double widthChange = (double)windowRect.Width / this.ResolutionWidth;
            double heightChange = (double)windowRect.Height / this.ResolutionHeight;
            return new Rectangle(
                     x: (int)Math.Round(rect.Left * widthChange),
                     y: (int)Math.Round(rect.Top * heightChange),
                 width: (int)Math.Round(rect.Width * widthChange),
                height: (int)Math.Round(rect.Height * heightChange));
        }

        /// <summary>
        /// Gets the background color for the trait icon
        /// </summary>
        /// <param name="iconRect">Trait rectangle</param>
        /// <param name="windowRect">Window rectangle</param>
        /// <returns>Background color</returns>
        private Color GetIconBackgroundColor(Rectangle iconRect, Rectangle windowRect)
        {
            Bitmap bitmap = new Bitmap(iconRect.Width, iconRect.Height);
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.CopyFromScreen(windowRect.Left + iconRect.Left, windowRect.Top + iconRect.Top, 0, 0, iconRect.Size);
            }

            bitmap = (Bitmap)bitmap.Resize(50, 50);
            return bitmap.GetPixel(5, 20);
        }

        /// <summary>
        /// Determines which traits are checked
        /// </summary>
        /// <param name="desiredTrait1">First desired trait</param>
        /// <param name="desiredTrait2">Second desired trait</param>
        private void GetDesiredTraits(ref Trait desiredTrait1, ref Trait desiredTrait2)
        {
            foreach (var checkbox in this.checkboxes)
            {
                if (checkbox.Checked)
                {
                    if (desiredTrait1 == Trait.Unknown)
                    {
                        desiredTrait1 = (Trait)checkbox.Tag;
                    }
                    else
                    {
                        desiredTrait2 = (Trait)checkbox.Tag;
                        return;
                    }
                }
            }
        }

        /// <summary>
        /// Moves the mouse to the center of the desired area on the Magicite window
        /// </summary>
        /// <param name="windowRect">Magicite window</param>
        /// <param name="rect">Desired area inside Magicite</param>
        private void MoveMouse(Rectangle windowRect, Rectangle rect)
        {
            int x = windowRect.Left + rect.Left + ((rect.Right - rect.Left) / 2);
            int y = windowRect.Top + rect.Top + ((rect.Bottom - rect.Top) / 2);
            Cursor.Position = new Point(x, y);
        }

        /// <summary>
        /// Clicks the mouse repeatedly at the desired location until the given condition is met 
        /// </summary>
        /// <param name="x">X coordinate</param>
        /// <param name="y">Y coordinate</param>
        /// <param name="stopCondition">Stopping condition</param>
        private void ClickMouse(int x, int y, Func<bool> stopCondition)
        {
            double maxSeconds = (double)numericUpDownMaxSeconds.Value;
            bool timeLimited = maxSeconds > 0;
            DateTime endTime = DateTime.Now.AddSeconds(maxSeconds);
            while (!stopCondition.Invoke() && (!timeLimited || DateTime.Now < endTime))
            {
                System.Threading.Thread.Sleep(1);
                Win32.mouse_event(Win32.MOUSEEVENTF_LEFTDOWN | Win32.MOUSEEVENTF_LEFTUP, (uint)x, (uint)y, 0, 0);
            }
        }
    }
}