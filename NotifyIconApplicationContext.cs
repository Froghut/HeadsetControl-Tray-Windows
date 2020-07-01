using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using Timer = System.Windows.Forms.Timer;

namespace HeadsetControl_Tray_Windows
{
	class NotifyIconApplicationContext : ApplicationContext
	{
		#region Enums

		private enum ChargeStatus
		{
			Disconnected,
			Charging,
			Discharging
		}

		#endregion

		#region Private Fields

		private NotifyIcon _notifyIcon;
		private Timer _timer;
		private Font _fontDiconnected;
		private Font _fontCharging;
		private Font _fontDischarging;
		private Bitmap _bitmap;
		private Graphics _graphics;
		private Regex _regex;
		private SolidBrush _brush;

		#endregion

		#region Constructors

		public NotifyIconApplicationContext()
		{
			_fontDiconnected = new Font("Microsoft Sans Serif", 24, FontStyle.Bold, GraphicsUnit.Pixel);
			_fontCharging = new Font("Microsoft Sans Serif", 16, FontStyle.Bold, GraphicsUnit.Pixel);
			_fontDischarging = new Font("Microsoft Sans Serif", 16, FontStyle.Regular, GraphicsUnit.Pixel);
			_bitmap = new Bitmap(16, 16);
			_brush = new SolidBrush(Color.White);
			_graphics = Graphics.FromImage(_bitmap);
			_graphics.TextRenderingHint = TextRenderingHint.SingleBitPerPixelGridFit;
			_regex = new Regex("Battery: (.*?)[%\r]");

			MenuItem exitMenuItem = new MenuItem("Exit", new EventHandler(Exit));

			_notifyIcon = new NotifyIcon();
			_notifyIcon.ContextMenu = new ContextMenu(new MenuItem[] { exitMenuItem });
			UpdateIcon();
			_notifyIcon.Visible = true;

			_timer = new Timer();
			_timer.Interval = 5000;
			_timer.Tick += (s, e) => UpdateIcon();
			_timer.Start();
		}

		#endregion

		#region Private methods

		private void UpdateIcon()
		{
			try
			{
				string chargeStatusString;
				ChargeStatus chargeStatus = GetChargeStatus(out chargeStatusString);
				_graphics.Clear(Color.Transparent);
				switch (chargeStatus)
				{
					case ChargeStatus.Disconnected:
						_graphics.DrawString(chargeStatusString, _fontDiconnected, _brush, 0, -9);
						break;
					case ChargeStatus.Charging:
						//_graphics.DrawString(chargeStatusString, _fontCharging, _brush, 0, -2);
						TextRenderer.DrawText(_graphics, chargeStatusString, _fontCharging, new Point(-4, -3), Color.White);
						break;
					case ChargeStatus.Discharging:
						_graphics.DrawString(chargeStatusString, _fontDischarging, _brush, -4, -2);
						break;
				}

				IntPtr hIcon = _bitmap.GetHicon();
				_notifyIcon.Icon = Icon.FromHandle(hIcon);
			}
			catch (Exception ex)
			{
				ThreadPool.QueueUserWorkItem(o => MessageBox.Show(ex.ToString(), "HeadsetControl Tray"));
			}
		}

		private ChargeStatus GetChargeStatus(out string chargeStatusString)
		{
			ProcessStartInfo psi = new ProcessStartInfo("headsetcontrol.exe", "-b");
			psi.CreateNoWindow = true;
			psi.RedirectStandardOutput = true;
			psi.UseShellExecute = false;
			Process process = Process.Start(psi);
			string str = process.StandardOutput.ReadToEnd();
			process.WaitForExit();
			Match match = _regex.Match(str);
			if (match.Success)
			{
				string value = match.Groups[1].Value;
				if (value == "Charging")
				{
					chargeStatusString = "⚡";
					return ChargeStatus.Charging;
				}
				else
				{
					chargeStatusString = value.PadLeft(2, '0');
					if (chargeStatusString == "100")
						chargeStatusString = "99";
					return ChargeStatus.Discharging;
				}
			}

			chargeStatusString = "-";
			return ChargeStatus.Disconnected;
		}

		private void Exit(object sender, EventArgs e)
		{
			_notifyIcon.Visible = false;
			Application.Exit();
		}

		#endregion
	}
}