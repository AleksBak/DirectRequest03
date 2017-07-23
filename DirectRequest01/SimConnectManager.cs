using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using LockheedMartin.Prepar3D.SimConnect;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Threading;
using System.Runtime.ExceptionServices;
using System.Security;

namespace DirectRequest
{
	public enum UNITS
	{
		BOOL = 0,
		NUMBER,
		POSITION,
		PERCENT_OVER_100,

		KNOTS,
		FEET,
		FEET_PER_SECOND,
		FEET_PER_SECOND_SQUARED,

		KILOMETERS,
		METERS,
		METER_PER_SECOND,
		METER_PER_SECOND_SQUARED,

		RADIANS,
		DEGREES,

		ENUM,
	}

	public class DataField
	{
		public string DatumName { get; set; }
		public string UnitsName { get; set; }

		/// <summary>
		/// при получении данных помним: сюда значение пишет класс SimConnectManager, а мы только читаем :)
		/// </summary>
		public double Value { get; set; }

		public DataField(string datumName, string unitsName)
		{
			DatumName = datumName;
			UnitsName = unitsName;
		}

		public override bool Equals(object obj)
		{
			return ((DataField)obj).DatumName == this.DatumName && ((DataField)obj).UnitsName == this.UnitsName;
		}
	}

	/// <summary>
	/// Что нужно в этом классе:
	/// 0. метод 'ReceiveMessage' для обработки полученного сообщения <see cref="System.Windows.Forms"/> для формы приложения <see cref="ownerFormHandle"/> 
	/// 1. метод 'Connect' - подключаемся к P3D;
	/// 2. метод 'AddToDataDefinition' - добавляем соотв. поле запроса в нашу структуру запроса;
	/// 3. метод 'RegisterDataDefineStruct' - регистрируем нашу структуру запроса;
	/// 4. метод 'RequestData' - отправка запроса SimConnect-у для выссылки им запрошенных данных
	/// 5. метод 'CloseConnection' - чтобы отключиться от SimConnect-а
	/// 6. событие 'OnGetData' - которое возникает при получении ранее запрошенных данных;
	/// </summary>
	public class SimConnectManager : IDisposable
	{
		#region private члены

		/// <summary>
		/// Это дескриптор окна (HWND) приложения, с которым связан элемент управления.
		/// </summary>
		private IntPtr ownerFormHandle { get; }

		/// <summary>
		/// Кол-во полей в структуре запроса данных от SimConnect-а.
		/// Если на форме приложения <see cref="ownerFormHandle"/> будет другое кол-во, то в др. инф. полях формы ничего не отобразится.
		/// </summary>
		private const int FIELDS_COUNT = 23;

		private List<DataField> DataFields = new List<DataField>(FIELDS_COUNT);

		/// <summary>
		/// SimConnect object
		/// </summary>
		private SimConnect simconnect = null;

		/// <summary>
		/// User-defined win32 event
		/// </summary>
		private const int WM_USER_SIMCONNECT = 0x0402;

		enum DEFINITIONS
		{
			myDefineID,
		}

		enum DATA_REQUESTS
		{
			myRequestID,
		};

		/// <summary>
		/// This is how you declare a data structure so that simconnect knows how to fill it/read it.
		/// Если на форме приложения <see cref="ownerFormHandle"/> будет другое кол-во чем тут полей, то в др. инф. полях формы ничего не отобразится.
		/// </summary>
		[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi, Pack = 1)]
		private unsafe struct Struct1
		{
			public fixed double fieds[FIELDS_COUNT];
		};

		/// <summary>
		/// Если смогли подключиться к SimConnect-у, то сразу усстанавливаем этот флаг.
		/// При обрывах связи или же при дисконнекте сбрасываем его.
		/// </summary>
		private bool isConnectedToP3D = false;

		#endregion

		public SimConnectManager(IntPtr ownerForm)
		{
			this.ownerFormHandle = ownerForm;
		}

		#region Обработчики событий

		/// <summary>
		/// Обработчик события подключения к SimConnect-у
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="data"></param>
		void simconnect_OnRecvOpen(SimConnect sender, SIMCONNECT_RECV_OPEN data)
		{
			this.isConnectedToP3D = true;

			// на всякий случай создаем копию события т.к. возможна ситуация, что подписчик отпишется от события в момент проверки ниже
			ConnectedEventHandler connectedHandler = ConnectedEvent;

			// далее проверяем, что есть ли какой-то подписчик на событие ConnectedToEndPointEvent (эта переменная не равна нулю) и генерируем тогда такое событие
			if (connectedHandler != null)
			{
				connectedHandler(this, new EventArgs());
			}
		}

		/// <summary>
		/// The case where the user closes Prepar3D
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="data"></param>
		void simconnect_OnRecvQuit(SimConnect sender, SIMCONNECT_RECV data)
		{
			CloseConnection();

			// на всякий случай создаем копию события т.к. возможна ситуация, что подписчик отпишется от события в момент проверки ниже
			DisconnectedEventHandler disconnectedHandler = DisconnectedEvent;

			// далее проверяем, что есть ли какой-то подписчик на событие ConnectedToEndPointEvent (эта переменная не равна нулю) и генерируем тогда такое событие
			if (disconnectedHandler != null)
			{
				disconnectedHandler(this, new EventArgs());
			}
		}

		void simconnect_OnRecvException(SimConnect sender, SIMCONNECT_RECV_EXCEPTION data)
		{
			Thread.Sleep(2);
			//displayText("Exception received: " + data.dwException);
		}

		/// <summary>
		/// Обработчик события прихода данных от SimConnect
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="data"></param>
		void simconnect_OnRecvSimobjectDataBytype(SimConnect sender, SIMCONNECT_RECV_SIMOBJECT_DATA_BYTYPE data)
		{
			switch ((DATA_REQUESTS)data.dwRequestID)
			{
				case DATA_REQUESTS.myRequestID:
					{
						unsafe
						{
							Struct1 s1 = (Struct1)data.dwData[0];
							int i = 0;
							foreach (DataField field in DataFields)
							{
								field.Value = s1.fieds[i++];
							}
						};

						// на всякий случай создаем копию события т.к. возможна ситуация, что подписчик отпишется от события в момент проверки ниже
						ReceivedDataEventHandler receivedDataHandler = ReceivedDataEvent;

						// далее проверяем, что есть ли какой-то подписчик на событие ConnectedToEndPointEvent (эта переменная не равна нулю) и генерируем тогда такое событие
						if (receivedDataHandler != null)
						{
							receivedDataHandler(this, new ReceivedDataEventArgs(this.DataFields));
						}
					}
					break;

				default:
					{
						// на всякий случай создаем копию события т.к. возможна ситуация, что подписчик отпишется от события в момент проверки ниже
						UnknownRequestIDEventHandler unknownRequestIDHandler = UnknownRequestIDEvent;

						// далее проверяем, что есть ли какой-то подписчик на событие ConnectedToEndPointEvent (эта переменная не равна нулю) и генерируем тогда такое событие
						if (unknownRequestIDHandler != null)
						{
							unknownRequestIDHandler(this, new UnknownRequestIDEventArgs(data.dwRequestID));
						}
					}
					break;
			}
		}

		#endregion

		#region Методы

		/// <summary>
		/// метод 'ReceiveMessage' для обработки полученного сообщения <see cref="System.Windows.Forms"/> для формы приложения <see cref="ownerFormHandle"/>
		/// </summary>
		/// <param name="m"></param>
		/// <returns></returns>
		[HandleProcessCorruptedStateExceptions]
		[SecurityCritical]
		public bool ReceiveMessage(ref Message m)
		{
			if (m.Msg == WM_USER_SIMCONNECT && simconnect != null)
			{
				try
				{
					simconnect.ReceiveMessage();
					return true;
				}
				catch (System.AccessViolationException ex)
				//catch (Exception ex)
				{
					// на всякий случай создаем копию события т.к. возможна ситуация, что подписчик отпишется от события в момент проверки ниже
					DisconnectedEventHandler disconnectedHandler = DisconnectedEvent;

					// далее проверяем, что есть ли какой-то подписчик на событие ConnectedToEndPointEvent (эта переменная не равна нулю) и генерируем тогда такое событие
					if (disconnectedHandler != null)
					{
						disconnectedHandler(this, new EventArgs());
					}

					return false;
				}
			}
			else
			{
				return false;
			}
		}

		/// <summary>
		/// В этом методе просто подключаемся и потом ждем соотв. события от SimConnect-а.
		/// </summary>
		public void Connect()
		{
			if (simconnect == null)
			{
				try
				{
					// the constructor is similar to SimConnect_Open in the native API
					simconnect = new SimConnect("Managed Data Request", this.ownerFormHandle, WM_USER_SIMCONNECT, null, 0);

					// listen to connect, quit msgs and to exceptions
					simconnect.OnRecvOpen += new SimConnect.RecvOpenEventHandler(simconnect_OnRecvOpen);
					simconnect.OnRecvQuit += new SimConnect.RecvQuitEventHandler(simconnect_OnRecvQuit);
					simconnect.OnRecvException += new SimConnect.RecvExceptionEventHandler(simconnect_OnRecvException);
				}
				catch (COMException ex)
				{
					//displayText("Unable to connect to Prepar3D:\n\n" + ex.Message);
				}
			}
			else
			{
				//displayText("Error - try again");
				//closeConnection();
				//setButtons(true, false, false);
			}
		}

		/// <summary>
		/// добавляем соотв. поле запроса в нашу структуру запроса
		/// </summary>
		/// <param name="datumName"></param>
		/// <param name="unitsName"></param>
		public void AddToDataDefinition(string datumName, string unitsName)
		{
			try
			{
				if (Enum.GetNames(typeof(UNITS)).Contains(unitsName))// && DataFields.IndexOf(DataFields.Last()) < FIELDS_COUNT)
				{
					// т.к. в перичислении не используем имена полей с пробелами, а с '_', то меняем если такие есть на пробелы:
					simconnect.AddToDataDefinition(DEFINITIONS.myDefineID, datumName, unitsName.Replace('_', ' '), SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);
					DataFields.Add(new DataField(datumName, unitsName));
				}
			}
			catch (COMException ex)
			{
				Thread.Sleep(2);
				//displayText(ex.Message);
			}
		}

		/// <summary>
		/// IMPORTANT: register it with the simconnect managed wrapper marshaller, if you skip this step, you will only receive a uint in the .dwData field.
		/// Этот метод вызываем в конце, после добавлений полей для структуры запроса к SimConnect-у <see cref="AddToDataDefinition"/>.
		/// </summary>
		public void RegisterDataDefineStruct()
		{
			try
			{
				// IMPORTANT: register it with the simconnect managed wrapper marshaller, if you skip this step, you will only receive a uint in the .dwData field.
				simconnect.RegisterDataDefineStruct<Struct1>(DEFINITIONS.myDefineID);

				// catch a simobject data request
				simconnect.OnRecvSimobjectDataBytype += new SimConnect.RecvSimobjectDataBytypeEventHandler(simconnect_OnRecvSimobjectDataBytype);
			}
			catch (COMException ex)
			{
				Thread.Sleep(2);
				//displayText(ex.Message);
			}
		}

		/// <summary>
		/// Отправка запроса SimConnect-у для выссылки им ранее запрошенных данных
		/// </summary>
		public void RequestData()
		{
			if (simconnect != null)
			{
				try
				{
					// The following call returns identical information to:
					// simconnect.RequestDataOnSimObject(DATA_REQUESTS.REQUEST_1, DEFINITIONS.Struct1, SimConnect.SIMCONNECT_OBJECT_ID_USER, SIMCONNECT_PERIOD.ONCE);

					simconnect.RequestDataOnSimObjectType(DATA_REQUESTS.myRequestID, DEFINITIONS.myDefineID, 0, SIMCONNECT_SIMOBJECT_TYPE.USER);
					//displayText("Request sent...");
				}
				catch (COMException ex)
				{
					Thread.Sleep(2);
					//displayText(ex.Message);
				}
			}
		}

		/// <summary>
		/// Отключение от SimConnect-а. По заершении работы следует также обязательно отключиться от SimConnect-а через этот метод.
		/// </summary>
		public void CloseConnection()
		{
			if (simconnect != null)
			{
				// Dispose serves the same purpose as SimConnect_Close()
				simconnect.Dispose();
				simconnect = null;
				this.isConnectedToP3D = false;
			}
		}

		public void Dispose()
		{
			if (simconnect != null)
			{
				// Dispose serves the same purpose as SimConnect_Close()
				((IDisposable)simconnect).Dispose();
				simconnect = null;
				//displayText("Connection closed");
			}
		}

		public bool IsConnected()
		{
			return this.isConnectedToP3D;
		}

		#endregion

		#region Делегаты, генерация событий и пр.

		/// <summary>
		/// Класс является издателем события об подключении к конечной точке. Этот делегат нужен для генерации события.
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		public delegate void ConnectedEventHandler(object sender, EventArgs e);

		/// <summary>
		/// Событие в этом классе публикации, что наш клиент подключился к конечной точке.
		/// </summary>
		public event ConnectedEventHandler ConnectedEvent;


		/// <summary>
		/// Пользовательский класс EventArgs для события получения данных от конечной точки (SimConnect-а).
		/// </summary>
		public class ReceivedDataEventArgs : EventArgs
		{
			public ReceivedDataEventArgs(List<DataField> dataFields)
			{
				this.DataFields = dataFields;
			}

			/// <summary>
			/// Готовый список полученных данных от SimConnect-а
			/// </summary>
			public List<DataField> DataFields;
		}

		/// <summary>
		/// Класс является издателем события об поолучении данных от конечной точки. Этот делегат нужен для генерации события.
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		public delegate void ReceivedDataEventHandler(object sender, ReceivedDataEventArgs e);

		/// <summary>
		/// Событие в этом классе публикации, что наш клиент поолучил данные от конечной точки.
		/// </summary>
		public event ReceivedDataEventHandler ReceivedDataEvent;


		/// <summary>
		/// Класс является издателем события об отключении от конечной точки. Этот делегат нужен для генерации события.
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		public delegate void DisconnectedEventHandler(object sender, EventArgs e);

		/// <summary>
		/// Событие в этом классе публикации, что наш клиент отключился от конечной точки.
		/// </summary>
		public event DisconnectedEventHandler DisconnectedEvent;


		/// <summary>
		/// Пользовательский класс EventArgs для данных события "Unknown request ID".
		/// </summary>
		public class UnknownRequestIDEventArgs : EventArgs
		{
			public UnknownRequestIDEventArgs(uint dwRequestID)
			{
				this.dwRequestID = dwRequestID;
			}

			/// <summary>
			/// Received Request ID
			/// </summary>
			public uint dwRequestID { get; }
		}

		/// <summary>
		/// Класс является издателем события об поолучении "Unknown request ID" от конечной точки. Этот делегат нужен для генерации события.
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		public delegate void UnknownRequestIDEventHandler(object sender, UnknownRequestIDEventArgs e);

		/// <summary>
		/// Событие в этом классе публикации, что наш клиент поолучил "Unknown request ID" от конечной точки.
		/// </summary>
		public event UnknownRequestIDEventHandler UnknownRequestIDEvent;

		#endregion
	}
} 
