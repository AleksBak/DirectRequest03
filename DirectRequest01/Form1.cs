using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace DirectRequest
{
	public partial class Form1 : Form
	{
		SimConnectManager simConnectManager = null;
		Stopwatch sw = new Stopwatch();

		Math3D.Cube mainCube;
		Point drawOrigin;

		#region Основные методы

		public Form1()
		{
			InitializeComponent();

			// инициализируем переменную делегата для вызова метода вывода информации об скважине на форме
			PrintDataOnFormFunc = new PrintDataOnForm(FillControlsAndProcessingData);
		}

		private void Form1_Load(object sender, EventArgs e)
		{
			mainCube = new Math3D.Cube(150, 7, 150);
			drawOrigin = new Point(pictureBox1.Width / 2, pictureBox1.Height / 2);
			mainCube.DrawWires = true;
			mainCube.FillFront = true;
			mainCube.FillBack = true;
			mainCube.FillLeft = true;
			mainCube.FillRight = true;
			mainCube.FillTop = true;
			mainCube.FillBottom = true;
		}

		private void Form1_Paint(object sender, PaintEventArgs e)
		{
			Render();
		}

		protected override void DefWndProc(ref Message m)
		{
			if (simConnectManager == null || simConnectManager.ReceiveMessage(ref m) == false)
			{
				base.DefWndProc(ref m);
			}
		}

		private void Disconnect()
		{
			if (simConnectManager != null)
			{
				MainTimer.Enabled = false;
				StopButton.Enabled = false;

				simConnectManager.CloseConnection();

				simConnectManager.ConnectedEvent -= SimConnectManager_ConnectedEvent;
				simConnectManager.DisconnectedEvent -= SimConnectManager_DisconnectedEvent;
				simConnectManager.UnknownRequestIDEvent -= SimConnectManager_UnknownRequestIDEvent;
				simConnectManager.ReceivedDataEvent -= SimConnectManager_ReceivedDataEvent;

				simConnectManager.Dispose();
				simConnectManager = null;

				displayText("Connection closed");
				sw.Stop();
			}
		}

		private void AddDataDefinitionsFromFormAndRegister()
		{
			foreach (Control control in this.Controls)
			{
				if (control is GroupBox)
				{
					foreach (Control ctl in control.Controls)
					{
						if (ctl is Label && ((ctl is LinkLabel) == false && ctl.Tag != null && ctl.Tag.ToString().Length > 0))
						{
							simConnectManager.AddToDataDefinition(ctl.Text, ctl.Tag.ToString());
						}
					}
				}
			}
			simConnectManager.RegisterDataDefineStruct();
		}

		private void Connect()
		{
			if (simConnectManager == null)
			{
				simConnectManager = new SimConnectManager(this.Handle);
			}

			simConnectManager.ConnectedEvent += SimConnectManager_ConnectedEvent;
			simConnectManager.DisconnectedEvent += SimConnectManager_DisconnectedEvent;
			simConnectManager.UnknownRequestIDEvent += SimConnectManager_UnknownRequestIDEvent;
			simConnectManager.ReceivedDataEvent += SimConnectManager_ReceivedDataEvent;

			simConnectManager.Connect();
			sw.Start();
		}

		private void Render()
		{
			mainCube.RotateX = (float)pitchBasicProgressBar.Value;
			//mainCube.RotateY = 0;		// (float)headingBasicProgressBar.Value;
			mainCube.RotateZ = (float)bankBasicProgressBar.Value;

			pictureBox1.Image = mainCube.DrawCube(drawOrigin);
		}

		long ticksCount = 0;
		long maxLatency = 0;

		/// <summary>
		/// Метод заполнения компонентов на форме значениями от принятых SimConnect данных, а также их дальнейшей обработки
		/// </summary>
		/// <param name="dataFields"></param>
		private void FillControlsAndProcessingData(List<DataField> dataFields)
		{
			long currTicks = sw.ElapsedTicks;
			long currLatency = sw.ElapsedMilliseconds;
			if (maxLatency < currLatency)
			{
				maxLatency = currLatency;
			}
			LatencyBasicProgressBar.Value = currLatency;
			LatencyLabel.Text = currLatency.ToString() + " (max " + maxLatency.ToString() + ")";

			CountLabel.Text = (++ticksCount).ToString();

			foreach (Control control in this.Controls)
			{
				if (control is GroupBox)
				{
					DataField field = null;
					bool fieldFound = false;

					// ищем среди всех значений полученного списка значений соотв. данному GroupBox текущее значение:
					foreach (Control ctl in control.Controls)
					{
						if (ctl is Label && ((ctl is LinkLabel) == false && ctl.Tag != null && ctl.Tag.ToString().Length > 0))
						{
							field = new DataField(ctl.Text, ctl.Tag.ToString());
							if (dataFields.Contains(field) == true)
							{
								field.Value = dataFields[dataFields.IndexOf(field)].Value;
								fieldFound = true;
							}
						}
					}

					// отображаем текущее значение в BasicProgressBar и если надо переносим в MIN/MAX значеения (в соотв. LinkLabel - ах):
					if (fieldFound == true && field != null)
					{
						foreach (Control ctl in control.Controls)
						{
							if (ctl is BasicProgressBar)
							{
								(ctl as BasicProgressBar).Value = (float)Math.Round(field.Value, 2, MidpointRounding.AwayFromZero);
							}

							if (ctl is LinkLabel)
							{
								if (ctl.Tag.ToString() == "MIN")
								{
									double val = 0;
									Double.TryParse(ctl.Text, out val);
									if (val > field.Value)
									{
										ctl.Text = Math.Round(field.Value, 2, MidpointRounding.AwayFromZero).ToString();
									}
								}
								else if (ctl.Tag.ToString() == "MAX")
								{
									double val = 0;
									Double.TryParse(ctl.Text, out val);
									if (val < field.Value)
									{
										ctl.Text = Math.Round(field.Value, 2, MidpointRounding.AwayFromZero).ToString();
									}
								}
							}
						}
					}
				}
			}

			Render();

			TotalTimeLabel.Text = (sw.ElapsedMilliseconds - currLatency).ToString();
		}

		#endregion

		#region Обработчики событий

		/// <summary>
		/// Объявление делегата для метода заполнения компонентов на форме (панели и соотв. GroupBox) значениями от принятых SimConnect данных
		/// </summary>
		/// <param name="datafields"></param>
		delegate void PrintDataOnForm(List<DataField> datafields);

		/// <summary>
		/// Переменная для вышеобъявленного делегата метода заполнения компонентов на форме значениями от принятых SimConnect данных
		/// </summary>
		private PrintDataOnForm PrintDataOnFormFunc;

		private void SimConnectManager_ReceivedDataEvent(object sender, SimConnectManager.ReceivedDataEventArgs e)
		{
			try
			{
				// вызываем метод вывода данных в компонент на форме, созданный в другом потоке, через предназначенный для этого метод Invoke
				this.Invoke(PrintDataOnFormFunc, new object[] { e.DataFields });
			}
			// если форма закрылась уже (или после внезапного спящего режима и т.п.):
			catch (ObjectDisposedException ex)
			{
				// ..
				Thread.Sleep(1);
			}
			catch (Exception ex)
			{
				// ..
				Thread.Sleep(1);
			}
		}

		private void SimConnectManager_UnknownRequestIDEvent(object sender, SimConnectManager.UnknownRequestIDEventArgs e)
		{
			displayText("Unknown request ID: " + e.dwRequestID);
		}

		private void SimConnectManager_DisconnectedEvent(object sender, EventArgs e)
		{
			MainTimer.Enabled = false;
			StopButton.Enabled = false;
			displayText("Prepar3D has exited");
			Disconnect();
		}

		private void SimConnectManager_ConnectedEvent(object sender, EventArgs e)
		{
			displayText("Connected to Prepar3D");
			AddDataDefinitionsFromFormAndRegister();
			StopButton.Enabled = true;
			MainTimer.Enabled = true;
		}

		#endregion

		#region Common 
		private void Form1_FormClosing(object sender, FormClosingEventArgs e)
		{
			MainTimer.Enabled = false;
			Disconnect();
		}

		private void StartButton_Click(object sender, EventArgs e)
		{
			Connect();
		}

		private void StopButton_Click(object sender, EventArgs e)
		{
			Disconnect();
		}

		private void MainTimer_Tick(object sender, EventArgs e)
		{
			if (simConnectManager != null && simConnectManager.IsConnected() == true)
			{
				sw.Restart();
				simConnectManager.RequestData();
				tickRadioButton.Checked = !tickRadioButton.Checked;
			}
		}

		// Response number
		int response = 1;

		// Output text - display a maximum of 3 lines
		string output = "\n\n";

		private void ResetButton_Click(object sender, EventArgs e)
		{

		}

		void displayText(string s)
		{
			// remove first string from output
			output = output.Substring(output.IndexOf("\n") + 1);

			// add the new string
			output += "\n" + response++ + ": " + s;

			// display it
			InfoRichTextBox.Text = output;
		}

		private void toolTip1_Popup(object sender, PopupEventArgs e)
		{
			ToolTip tt = sender as ToolTip;

			Control ctl = e.AssociatedControl;
			if (ctl is LinkLabel || (ctl is BasicProgressBar && ctl.Tag == null))
			{
				GroupBox gbx = ctl.Parent as GroupBox;
				if (gbx != null)
				{
					foreach (Control ctlLabel in gbx.Controls)
					{
						if (ctlLabel is Label && (ctlLabel is LinkLabel) == false)
						{
							tt.ToolTipTitle = ctlLabel.Text;
							break;
						}
					}
				}
			}
			else
			{
				tt.ToolTipTitle = (ctl.Tag == null) ? "Параметр" : ctl.Tag.ToString().Replace('_', ' ');
			}
		}

		#endregion
	}
}
