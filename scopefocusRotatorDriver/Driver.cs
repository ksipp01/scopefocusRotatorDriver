//tabs=4
// --------------------------------------------------------------------------------
// TODO fill in this information for your driver, then remove this line!
//
// ASCOM Rotator driver for scopefocus
//
// Description:	Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam 
//				nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam 
//				erat, sed diam voluptua. At vero eos et accusam et justo duo 
//				dolores et ea rebum. Stet clita kasd gubergren, no sea takimata 
//				sanctus est Lorem ipsum dolor sit amet.
//
// Implements:	ASCOM Rotator interface version: <To be completed by driver developer>
// Author:		(XXX) Your N. Here <your@email.here>
//
// Edit Log:
//
// Date			Who	Vers	Description
// -----------	---	-----	-------------------------------------------------------
// dd-mmm-yyyy	XXX	6.0.0	Initial edit, created from ASCOM driver template
// --------------------------------------------------------------------------------
//


// This is used to define code in the template that is specific to one class implementation
// unused code canbe deleted and this definition removed.
#define Rotator

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Runtime.InteropServices;

using ASCOM;
using ASCOM.Astrometry;
using ASCOM.Astrometry.AstroUtils;
using ASCOM.Utilities;
using ASCOM.DeviceInterface;
using System.Globalization;
using System.Collections;

namespace ASCOM.scopefocus
{
    //Started on 2-6-17
    // Your driver's DeviceID is ASCOM.scopefocus.Rotator
    //
    // The Guid attribute sets the CLSID for ASCOM.scopefocus.Rotator
    // The ClassInterface/None addribute prevents an empty interface called
    // _scopefocus from being created and used as the [default] interface
    //
    // TODO Replace the not implemented exceptions with code to implement the function or
    // throw the appropriate ASCOM exception.
    //

    /// <summary>
    /// ASCOM Rotator Driver for scopefocus.
    /// </summary>
    [Guid("60dd4181-dd94-4240-865b-ce0a49edb383")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Rotator : IRotatorV2
    {

        //****** add
        //  private Config config = new Config();
        private Serial serialPort;


        //  private TextWriter log;
        System.Threading.Mutex mutex = new System.Threading.Mutex();


        float lastPos = 0;
        //    double lastTemp = 0;
        bool lastMoving = false;
        bool lastLink = false;

        long UPDATETICKS = (long)(1 * 10000000.0); // 10,000,000 ticks in 1 second
        long lastUpdate = 0;


        long lastL = 0;
        //************end add




        /// <summary>
        /// ASCOM DeviceID (COM ProgID) for this driver.
        /// The DeviceID is used by ASCOM applications to load the driver at runtime.
        /// </summary>
        internal static string driverID = "ASCOM.scopefocus.Rotator";
        // TODO Change the descriptive string for your driver then remove this line
        /// <summary>
        /// Driver description that displays in the ASCOM Chooser.
        /// </summary>
        private static string driverDescription = "ASCOM Rotator Driver for scopefocus.";

        internal static string comPortProfileName = "COM Port"; // Constants used for Profile persistence
        internal static string comPortDefault = "COM1";
        internal static string traceStateProfileName = "Trace Level";
        internal static string traceStateDefault = "false";

        internal static string comPort; // Variables to hold the currrent device configuration
        internal static bool traceState;
        internal static int stepsPerDegree;

        /// <summary>
        /// Private variable to hold the connected state
        /// </summary>
        private bool connectedState;

        /// <summary>
        /// Private variable to hold an ASCOM Utilities object
        /// </summary>
        private Util utilities;

        /// <summary>
        /// Private variable to hold an ASCOM AstroUtilities object to provide the Range method
        /// </summary>
        private AstroUtils astroUtilities;

        /// <summary>
        /// Private variable to hold the trace logger object (creates a diagnostic log file with information that you specify)
        /// </summary>
        private TraceLogger tl;

        /// <summary>
        /// Initializes a new instance of the <see cref="scopefocus"/> class.
        /// Must be public for COM registration.
        /// </summary>
        public Rotator()
        {
            ReadProfile(); // Read device configuration from the ASCOM Profile store

            tl = new TraceLogger("", "scopefocus");
            tl.Enabled = traceState;
            tl.LogMessage("Rotator", "Starting initialisation");

            connectedState = false; // Initialise connected to false
            utilities = new Util(); //Initialise util object
            astroUtilities = new AstroUtils(); // Initialise astro utilities object
            //TODO: Implement your additional construction here

            tl.LogMessage("Rotator", "Completed initialisation");
        }


        //
        // PUBLIC COM INTERFACE IRotatorV2 IMPLEMENTATION
        //

        #region Common properties and methods.

        /// <summary>
        /// Displays the Setup Dialog form.
        /// If the user clicks the OK button to dismiss the form, then
        /// the new settings are saved, otherwise the old values are reloaded.
        /// THIS IS THE ONLY PLACE WHERE SHOWING USER INTERFACE IS ALLOWED!
        /// </summary>
        public void SetupDialog()
        {
            // consider only showing the setup dialog if not connected
            // or call a different dialog if connected
            if (IsConnected)
                System.Windows.Forms.MessageBox.Show("Already connected, just press OK");

            using (SetupDialogForm F = new SetupDialogForm())
            {
                var result = F.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    WriteProfile(); // Persist device configuration values to the ASCOM Profile store
                }
            }
        }

        public ArrayList SupportedActions
        {
            get
            {
                tl.LogMessage("SupportedActions Get", "Returning empty arraylist");
                return new ArrayList();
            }
        }

        public string Action(string actionName, string actionParameters)
        {
            if (actionName == "Home")
            {
                CommandString("H#", false);
                return "";
            }
            else
                return "";
          //  throw new ASCOM.ActionNotImplementedException("Action " + actionName + " is not implemented by this driver");
        }

        public void CommandBlind(string command, bool raw)
        {
            CheckConnected("CommandBlind");
            // Call CommandString and return as soon as it finishes
            this.CommandString(command, raw);
            // or
            throw new ASCOM.MethodNotImplementedException("CommandBlind");
            // DO NOT have both these sections!  One or the other
        }

        public bool CommandBool(string command, bool raw)
        {
            CheckConnected("CommandBool");
            string ret = CommandString(command, raw);
            // TODO decode the return string and return true or false
            // or
            throw new ASCOM.MethodNotImplementedException("CommandBool");
            // DO NOT have both these sections!  One or the other
        }

        public string CommandString(string command, bool raw)
        {
            CheckConnected("CommandString");
            // it's a good idea to put all the low level communication with the device here,
            // then all communication calls this function
            // you need something to ensure that only one command is in progress at a time

            //throw new ASCOM.MethodNotImplementedException("CommandString");

            if (!this.Connected)
            {
                throw new ASCOM.NotConnectedException();

            }

            string temp = "999";
            mutex.WaitOne();
            try
            {
                tl.LogMessage("Sending Command: ", command);
                if (!command.EndsWith("#"))
                    command += "#";


                serialPort.ClearBuffers();


                serialPort.Transmit(command);


                // get the return value
                temp = serialPort.ReceiveTerminated("#");


                serialPort.ClearBuffers();


                tl.LogMessage("Got Response: ", temp);

            }
            catch (Exception e)
            {
                tl.LogMessage("Caught exception in CommandString ", e.Message);

            }
            finally
            {
                mutex.ReleaseMutex();
            }

            return temp;
        }

        public void Dispose()
        {
            // Clean up the tracelogger and util objects
            tl.Enabled = false;
            tl.Dispose();
            tl = null;
            utilities.Dispose();
            utilities = null;
            astroUtilities.Dispose();
            astroUtilities = null;

            //**** added
            if (serialPort == null)
                return;
            serialPort.Connected = false;
            serialPort.Dispose();
            serialPort = null;
            //**** end addt
        }

        public bool Connected
        {
            get
            {
                tl.LogMessage("Connected Get", IsConnected.ToString());
                return IsConnected;
            }
            set
            {
                tl.LogMessage("Connected Set", value.ToString());
                if (value == IsConnected)
                    return;

                if (value)
                {
                    connectedState = true;
                    tl.LogMessage("Connected Set", "Connecting to port " + comPort);
                    // TODO connect to the device

                    // add

                    bool homeSet = false;
                    float posValue = 0;
                    bool setPos = false;
                 //   bool reverse = true;
                    bool contHold = false;
                    
                    // check if we are connected, return if we are
                    if (serialPort != null && serialPort.Connected)
                        return;
                    // get the port name from the profile
                    string portName;
                    using (ASCOM.Utilities.Profile p = new Profile())
                    {
                        // get the values that are stored in the ASCOM Profile for this driver
                        // these were usually set in the settings dialog
                        p.DeviceType = "Rotator";
                        if (!p.IsRegistered("ASCOM.AWR.Telescope"))  // added 2-28-16
                        {
                            p.Register("ASCOM.scopefocus.Rotator", "ASCOM Rotator Driver for scopefocus");
                        }

                //        homeSet = p.GetValue(driverID, "HomeSet").ToLower().Equals("true") ? true : false;
                        portName = p.GetValue(driverID, "ComPort");
                        //    portName = "COM4";
                        setPos = p.GetValue(driverID, "SetPos").ToLower().Equals("true") ? true : false;

                        // 6-16-16 added 2 lines below
                    //    reverse = p.GetValue(driverID, "Reverse").ToLower().Equals("true") ? true : false;
                        contHold = p.GetValue(driverID, "ContHold").ToLower().Equals("true") ? true : false;

                        if (setPos)
                            posValue = System.Convert.ToSingle(p.GetValue(driverID, "Pos"));
                   //     tempDisplay = p.GetValue(driverID, "TempDisp");
                        stepsPerDegree = Convert.ToInt32(p.GetValue(driverID, "StepsPerDegree"));
                        //blValue = System.Convert.ToInt32(p.GetValue(driverId, "BackLight"));

                        //*****temp rem until config is finished************


                        if (string.IsNullOrEmpty(portName))
                        {
                            // report a problem with the port name
                            throw new ASCOM.NotConnectedException("no Com port selected");
                        }

                        //*** end temp rem



                        // try to connect using the port
                        try
                        {
                            //    log = new StreamWriter("c:\\log.txt");
                            tl.LogMessage("Connecting to serial port", "");

                            // setup the serial port.

                            serialPort = new Serial();
                            serialPort.PortName = portName;
                            //   serialPort.PortName = comPort;
                            serialPort.Speed = SerialSpeed.ps9600;
                            serialPort.StopBits = SerialStopBits.One;
                            serialPort.Parity = SerialParity.None;
                            serialPort.DataBits = 8;
                            serialPort.DTREnable = false;


                            if (!serialPort.Connected)
                                serialPort.Connected = true;


                            // flush whatever is there.
                            serialPort.ClearBuffers();


                            // wait for the Serial Port to come online...better way to do this???
                            System.Threading.Thread.Sleep(1000);


                            // if the user is setting a position in the Settings dialog set it here.
                            if (setPos)
                                CommandString("P " + posValue * stepsPerDegree + "#", false);  //orig was M changed to P 10-18-2015 (want it to set the value not move)
                            //3-7-17 above also need to correct for user defined steps / degree (not just 100); 

                            // added 6-16-16 
                            //if (reverse)
                            //    CommandString("R 1#", false);
                            //else
                            //    CommandString("R 0#", false); // motor sitting shaft up turns clockwise with increasing numbers if NOT reversed

                            if (contHold)
                                CommandString("C 1#", false); //continuous hold on
                            else
                                CommandString("C 0#", false);


                            //   char td = tempDisplay.Length > 0 ? tempDisplay.ToUpper().ToCharArray()[0] : 'C';
                            //    CommandString("a" + td + "$", false);
                            //    SetRpm(System.Convert.ToInt32(p.GetValue(driverId, "RPM")));
                            // **** ADDED 2-28-16 ****
                            //turn off serialTrace if driverTrace is on.  
                            utilities = new Util(); //Initialise util object
                            if (traceState) // 6-17-16 changed from (!tracestate)
                                utilities.SerialTrace = false;
                            else
                                utilities.SerialTrace = true;



                        }
                        catch (Exception ex)
                        {
                            // report any error
                            throw new ASCOM.NotConnectedException("Serial port connectionerror", ex);
                        }
                    }




                    //  connectedState = true;
                    //  tl.LogMessage("Connected Set", "Connecting to port " + comPort);
                    // TODO connect to the device
                }
                else
                {
                    CommandString("C 0#", false); //release the continuous hold
                    System.Threading.Thread.Sleep(500);
                    //  Dispose();
                    connectedState = false;
                    tl.LogMessage("Connected Set", "Disconnecting from port " + comPort);
                    if (serialPort != null && serialPort.Connected)
                    {
                        //       CommandString("C 0#", false); //release the continuous hold
                        //       System.Threading.Thread.Sleep(500);
                        serialPort.Connected = false;
                        serialPort.Dispose();
                        serialPort = null;
                    }



                }
            }
        }


  //  }
                //else
                //{
                //    connectedState = false;
                //    tl.LogMessage("Connected Set", "Disconnecting from port " + comPort);
                //    // TODO disconnect from the device
                //}
         //   }
      //  }

        public string Description
        {
            // TODO customise this device description
            get
            {
                tl.LogMessage("Description Get", driverDescription);
                return driverDescription;
            }
        }

        public string DriverInfo
        {
            get
            {
                Version version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                // TODO customise this driver description
                string driverInfo = "Information about the driver itself. Version: " + String.Format(CultureInfo.InvariantCulture, "{0}.{1}", version.Major, version.Minor);
                tl.LogMessage("DriverInfo Get", driverInfo);
                return driverInfo;
            }
        }

        public string DriverVersion
        {
            get
            {
                Version version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                string driverVersion = String.Format(CultureInfo.InvariantCulture, "{0}.{1}", version.Major, version.Minor);
                tl.LogMessage("DriverVersion Get", driverVersion);
                return driverVersion;
            }
        }

        public short InterfaceVersion
        {
            // set by the driver wizard
            get
            {
                tl.LogMessage("InterfaceVersion Get", "2");
                return Convert.ToInt16("2");
            }
        }

        public string Name
        {
            get
            {
                string name = "scopefocus";
                tl.LogMessage("Name Get", name);
                return name;
            }
        }

        #endregion

        #region IRotator Implementation

        private float rotatorPosition = 0; // Absolute position angle of the rotator 

        public bool CanReverse
        {
            get
            {
                tl.LogMessage("CanReverse Get", false.ToString());
                return false;
            }
        }

        public void Halt()
        {

            CommandString("S#", false);  
          //  tl.LogMessage("Halt", "Not implemented");
          //  throw new ASCOM.MethodNotImplementedException("Halt");
        }

        public bool IsMoving
        {
            get
            {

                DoUpdate();
                return lastMoving;
                //tl.LogMessage("IsMoving Get", false.ToString()); // This rotator has instantaneous movement
                //return false;
            }
        }

        public bool Link
        {
            get
            {
                long now = DateTime.Now.Ticks;
                if (now - lastL > UPDATETICKS)
                {
                    if (serialPort != null)
                        lastLink = serialPort.Connected;

                    lastL = now;
                    return lastLink;
                }

                return lastLink;
            }
            set
            {
                this.Connected = value;
            }


            /*
            get
            {
                tl.LogMessage("Link Get", this.Connected.ToString());
                return this.Connected; // Direct function to the connected method, the Link method is just here for backwards compatibility
            }
            set
            {
                tl.LogMessage("Link Set", value.ToString());
                this.Connected = value; // Direct function to the connected method, the Link method is just here for backwards compatibility
            }
             */
        }


        public void Move(float Position)  // made int for testing...will need to have float or double 
        {  
            float moveTo = rotatorPosition * stepsPerDegree + Position * stepsPerDegree;  // corrects for 100 steps per degree, need to replace with user defined variable.  
            CommandString("M " + Math.Round(moveTo, 0) + "#", false);  // Position was 'int value' for focuser
            lastMoving = true;  //remd 1-12-15

          //  tl.LogMessage("Move", Position.ToString()); // Move by this amount
            //rotatorPosition += Position * stepsPerDegree;
            //rotatorPosition = (float)astroUtilities.Range(rotatorPosition, 0.0, true, 360.0, false); // Ensure value is in the range 0.0..359.9999...
        }

        public double PositionAngleToMotorSteps(float positionAngle)
        {
            var normalizedAngle = positionAngle % 360.0 * stepsPerDegree;
            return normalizedAngle;
        }

        public void MoveAbsolute(float Position)
        {
            var stepPosition = PositionAngleToMotorSteps(Position);

            CommandString("M " +  Math.Round(stepPosition, 0) + "#", false);  // Position was 'int value' for focuser  // corrects for 100 steps per degree, need to replace with user defined variable.  
            lastMoving = true;  //remd 1-12-15

       //     tl.LogMessage("MoveAbsolute", Position.ToString()); // Move to this position
            //rotatorPosition = Position * stepsPerDegree;
            //rotatorPosition = (float)astroUtilities.Range(rotatorPosition, 0.0, true, 360.0, false); // Ensure value is in the range 0.0..359.9999...
        }

        public float Position
        {
            get
            {
                DoUpdate();
                rotatorPosition = lastPos;
                return lastPos;

                //tl.LogMessage("Position Get", rotatorPosition.ToString()); // This rotator has instantaneous movement
                //return rotatorPosition;
            }
        }

        public bool Reverse
        {
            get
            {
                tl.LogMessage("Reverse Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("Reverse", false);
            }
            set
            {
                tl.LogMessage("Reverse Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("Reverse", true);
            }
        }

        public float StepSize
        {
            get
            {
                tl.LogMessage("StepSize Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("StepSize", false);
            }
        }

        public float TargetPosition
        {
            get
            {
                tl.LogMessage("TargetPosition Get", rotatorPosition.ToString()); // This rotator has instantaneous movement
                return rotatorPosition;
            }
        }

        private void DoUpdate()
        {
            // only allow access for "gets" once per second.
            // if inside of 1 second the buffered value will be used.
            if (DateTime.Now.Ticks > UPDATETICKS + lastUpdate)
            {
                lastUpdate = DateTime.Now.Ticks;


                // focuser returns a string like:
                // m:false;s:1000;t:25.20$
                //   m - denotes moving or not
                //   s - denotes the position in steps
                //   t - denotes the temperature, always in C


                String val = CommandString("G#", false);


                // split the values up.  Ideally you should check for null here.  
                // if something goes wrong this will throw an exception...no bueno...


                //focuser sends P 200;M true#  for e.g.

                String[] vals = val.Replace('#', ' ').Trim().Split(';');

                string valTrim = vals[0].Replace('#', ' ');
                string pos = valTrim.Replace('P', ' ').Trim();
                // these values are used in the "Get" calls.  That way the client gets an immediate
                // response.  However it may up to 1 second out of date.
                // Thus "lastMoving" must be set to true when the move is initiated in "Move"

                lastPos = Convert.ToSingle(pos)/100;  // correct for 100 steps per degree  ****need to fix to user defined variable.  ****
                //    lastMoving = false;
                lastMoving = vals[1].Substring(2) == "true" ? true : false;  //*** remd 1-12-15
                //   *** 1-12-15  to implement this need to change arduino code to retrun something liek "M:True" 
                //   *** like example above line 640, then slipt ther string into an array and decifer them



                //    lastPos = Convert.ToInt16(vals[1].Substring(2));
                //    lastTemp = Convert.ToDouble(vals[2].Substring(2));
            }
        }



        #endregion

        #region Private properties and methods
        // here are some useful properties and methods that can be used as required
        // to help with driver development

        #region ASCOM Registration

        // Register or unregister driver for ASCOM. This is harmless if already
        // registered or unregistered. 
        //
        /// <summary>
        /// Register or unregister the driver with the ASCOM Platform.
        /// This is harmless if the driver is already registered/unregistered.
        /// </summary>
        /// <param name="bRegister">If <c>true</c>, registers the driver, otherwise unregisters it.</param>
        private static void RegUnregASCOM(bool bRegister)
        {
            using (var P = new ASCOM.Utilities.Profile())
            {
                P.DeviceType = "Rotator";
                if (bRegister)
                {
                    P.Register(driverID, driverDescription);
                }
                else
                {
                    P.Unregister(driverID);
                }
            }
        }

        /// <summary>
        /// This function registers the driver with the ASCOM Chooser and
        /// is called automatically whenever this class is registered for COM Interop.
        /// </summary>
        /// <param name="t">Type of the class being registered, not used.</param>
        /// <remarks>
        /// This method typically runs in two distinct situations:
        /// <list type="numbered">
        /// <item>
        /// In Visual Studio, when the project is successfully built.
        /// For this to work correctly, the option <c>Register for COM Interop</c>
        /// must be enabled in the project settings.
        /// </item>
        /// <item>During setup, when the installer registers the assembly for COM Interop.</item>
        /// </list>
        /// This technique should mean that it is never necessary to manually register a driver with ASCOM.
        /// </remarks>
        [ComRegisterFunction]
        public static void RegisterASCOM(Type t)
        {
            RegUnregASCOM(true);
        }

        /// <summary>
        /// This function unregisters the driver from the ASCOM Chooser and
        /// is called automatically whenever this class is unregistered from COM Interop.
        /// </summary>
        /// <param name="t">Type of the class being registered, not used.</param>
        /// <remarks>
        /// This method typically runs in two distinct situations:
        /// <list type="numbered">
        /// <item>
        /// In Visual Studio, when the project is cleaned or prior to rebuilding.
        /// For this to work correctly, the option <c>Register for COM Interop</c>
        /// must be enabled in the project settings.
        /// </item>
        /// <item>During uninstall, when the installer unregisters the assembly from COM Interop.</item>
        /// </list>
        /// This technique should mean that it is never necessary to manually unregister a driver from ASCOM.
        /// </remarks>
        [ComUnregisterFunction]
        public static void UnregisterASCOM(Type t)
        {
            RegUnregASCOM(false);
        }

        #endregion

        /// <summary>
        /// Returns true if there is a valid connection to the driver hardware
        /// </summary>
        private bool IsConnected
        {
            get
            {
                // TODO check that the driver hardware connection exists and is connected to the hardware
                return connectedState;
            }
        }

        /// <summary>
        /// Use this function to throw an exception if we aren't connected to the hardware
        /// </summary>
        /// <param name="message"></param>
        private void CheckConnected(string message)
        {
            if (!IsConnected)
            {
                throw new ASCOM.NotConnectedException(message);
            }
        }

        /// <summary>
        /// Read the device configuration from the ASCOM Profile store
        /// </summary>
        internal void ReadProfile()
        {
            using (Profile driverProfile = new Profile())
            {
                driverProfile.DeviceType = "Rotator";
                traceState = Convert.ToBoolean(driverProfile.GetValue(driverID, traceStateProfileName, string.Empty, traceStateDefault));
                comPort = driverProfile.GetValue(driverID, comPortProfileName, string.Empty, comPortDefault);
            }
        }

        /// <summary>
        /// Write the device configuration to the  ASCOM  Profile store
        /// </summary>
        internal void WriteProfile()
        {
            using (Profile driverProfile = new Profile())
            {
                driverProfile.DeviceType = "Rotator";
                driverProfile.WriteValue(driverID, traceStateProfileName, traceState.ToString());
                driverProfile.WriteValue(driverID, comPortProfileName, comPort.ToString());
            }
        }

        #endregion

    }
}
