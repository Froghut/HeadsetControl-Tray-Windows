using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using HidLibrary;
using Timer = System.Windows.Forms.Timer;

namespace HeadsetControl_Tray_Windows
{
	class NotifyIconApplicationContext : ApplicationContext
	{
		#region Enums

		private const int _vendorId = 0x046d;
		private const int _productId = 0x0ab5;
		private bool _waitingForReport;


		[System.Runtime.InteropServices.DllImport("user32.dll", CharSet=CharSet.Auto)]
		static extern bool DestroyIcon(IntPtr handle);
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
		private SolidBrush _brushText;
		private SolidBrush _brushSidetone;
		private SolidBrush _brushCharging;
		private IntPtr? _hIcon = null;
		private HidDevice _device;
		private bool _sidetoneFull;
		private MenuItem _voltageMenuItem;
		private float _battery;
		private bool _charging;

		#endregion

		#region Constructors

		public NotifyIconApplicationContext()
		{
			_fontDiconnected = new Font("Microsoft Sans Serif", 24, FontStyle.Bold, GraphicsUnit.Pixel);
			_fontCharging = new Font("Microsoft Sans Serif", 16, FontStyle.Bold, GraphicsUnit.Pixel);
			_fontDischarging = new Font("Microsoft Sans Serif", 16, FontStyle.Regular, GraphicsUnit.Pixel);
			_bitmap = new Bitmap(16, 16);
			_brushText = new SolidBrush(Color.White);
			_brushSidetone = new SolidBrush(Color.LightGreen);
			_brushCharging = new SolidBrush(Color.FromArgb(255, 199, 96));
			_graphics = Graphics.FromImage(_bitmap);
			_graphics.TextRenderingHint = TextRenderingHint.SingleBitPerPixelGridFit;
			_regex = new Regex("Battery: (.*?)[%\r]");

			MenuItem exitMenuItem = new MenuItem("Exit", new EventHandler(Exit));
			_voltageMenuItem = new MenuItem("----");

			_notifyIcon = new NotifyIcon();
			_notifyIcon.ContextMenu = new ContextMenu(new MenuItem[] { _voltageMenuItem, exitMenuItem });
			_notifyIcon.MouseClick += NotifyIcon_OnMouseClick;
			_notifyIcon.DoubleClick += (sender, args) =>
			                           {
				                           UpdateUSB();
				                           _timer.Stop();
				                           _timer.Start();
			                           };
			_notifyIcon.Visible = true;


			List<HidDevice> hidDevices = new List<HidDevice>(HidDevices.Enumerate(_vendorId, _productId));
			if (hidDevices.Count == 4)
			{
				HidDevice device = hidDevices.First(d => d.Capabilities.Usage == 514);
				_device = device;
				_device.OpenDevice();
				Console.WriteLine(_device.IsOpen + " - " + _device.IsConnected);
			}


			_timer = new Timer();
			_timer.Interval = 30000;
			_timer.Tick += (s, e) => UpdateUSB();
			_timer.Start();
			UpdateUSB();
		}

		private void NotifyIcon_OnMouseClick(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Left)
			{
				byte[] data = new byte[] { 0x11, 0xff, 0x07, 0x1f, (byte) (_sidetoneFull ? 0x00 : 0x64), 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };

				HidReport hidReport = new HidReport(20, new HidDeviceData(data, HidDeviceData.ReadStatus.Success));


				_sidetoneFull = !_sidetoneFull;
				_device.WriteReport(hidReport);
				UpdateIcon();
			}
			else if (e.Button == MouseButtons.Middle)
			{
				UpdateUSB();
			}

		}

		#endregion

		private void UpdateUSB()
		{
			if (_waitingForReport)
				return;
			_waitingForReport = true;

			if (_device == null)
			{
				_waitingForReport = false;
				return;
			}

			try
			{
				byte[] data = new byte[20];
				data[0] = 0x11;
				data[1] = 0xff;
				data[2] = 0x08;
				data[3] = 0x0a;
				_device.Write(data);
				_device.ReadReport(Callback, 500);

			}
			catch (Exception ex)
			{
				Debug.WriteLine(ex);
				_waitingForReport = false;
				_battery = 0;
				UpdateIcon();
			}
		}

		public static string ByteArrayToString(byte[] ba)
		{
			StringBuilder hex = new StringBuilder(ba.Length * 2);
			foreach (byte b in ba)
				hex.AppendFormat("{0:x2}", b);
			return hex.ToString();
		}

		private void Callback(HidReport report)
		{
			_waitingForReport = false;

			if (report.Data[1] == 0x04)
			{
				_device.ReadReport(Callback);
				return;
			}

			int voltage = (report.Data[3] << 8) | report.Data[4];
			if (report.Data[1] == 0xff)
				voltage = 0;

			Debug.WriteLine(ByteArrayToString(report.Data));

			Debug.WriteLine(report.ReadStatus + " - " + voltage + " - " + report.Data.Length);

			_voltageMenuItem.Text = voltage.ToString();

			float battery = map(voltage, 3550, 4160, 1, 100);
			if (voltage == 0)
				battery = 0;
			if (battery < 0)
				battery = 0;
			if (battery > 99)
				battery = 99;
			_battery = battery;

			_charging = report.Data[5] == 0x03;

			UpdateIcon();
		}

		float map(float s, float from1, float from2, float to1, float to2)
		{
			return to1 + (s-from1)*(to2-to1)/(from2-from1);
		}

		#region Private methods

		private void UpdateIcon()
		{
			//try
			//{
			_waitingForReport = false;
			if (_hIcon != null)
				DestroyIcon(_hIcon.Value);
			string chargeStatusString;
			//ChargeStatus chargeStatus = GetChargeStatus(out chargeStatusString);
			_graphics.Clear(Color.Transparent);


			if (_battery == 0)
				_graphics.DrawString("-", _fontDiconnected, _brushText, 0, -9);
			else
			{
				string batteryString = _battery.ToString().PadLeft(2, '0');
				_graphics.DrawString(batteryString, _fontDischarging, _brushText, -4, -2);
				if (_sidetoneFull && _charging)
				{
					_graphics.DrawString(batteryString[0].ToString(), _fontDischarging, _brushCharging, -4f, -2);
					_graphics.DrawString(batteryString[1].ToString(), _fontDischarging, _brushSidetone, 5f, -2);
				}
				else if (_sidetoneFull)
				{
					_graphics.DrawString(batteryString[0].ToString(), _fontDischarging, _brushSidetone, -4f, -2);
					_graphics.DrawString(batteryString[1].ToString(), _fontDischarging, _brushSidetone, 5f, -2);

				}
				else if (_charging)
				{
					_graphics.DrawString(batteryString[0].ToString(), _fontDischarging, _brushCharging, -4f, -2);
					_graphics.DrawString(batteryString[1].ToString(), _fontDischarging, _brushCharging, 5f, -2);
				}

			}



			_hIcon = _bitmap.GetHicon();
			_notifyIcon.Icon = Icon.FromHandle(_hIcon.Value);
		}

		private void Exit(object sender, EventArgs e)
		{
			_notifyIcon.Visible = false;
			Application.Exit();
		}

		#endregion
	}
}