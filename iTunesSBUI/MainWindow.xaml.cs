using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using SharpBlade.Razer;
using SharpBlade.Native;
using iTunesLib;
using System.Runtime.InteropServices;
using System.IO;

namespace iTunesSBUI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        protected RazerManager _manager;
        protected iTunesApp _itunes;
        protected string _appPath;
        protected int _volume;
        protected bool _muted;

        const string _AUTHOR = "Tailor Made Solutions";
        const string _BRAND = "Razer";
        const string _APPNAME = "iTunes";

        public MainWindow()
        {
            InitializeComponent();
            try
            {
                _manager = RazerManager.Instance;
                _manager.AppEvent += OnAppEvent;

                createiTunes(true);

                setPlayPauseButton();
                RazerManager.Instance.EnableDynamicKey(RazerAPI.DynamicKeyType.DK6, dkBack, @"Default\Images\back.png", @"Default\Images\back_pressed.png", true);
                RazerManager.Instance.EnableDynamicKey(RazerAPI.DynamicKeyType.DK8, dkNext, @"Default\Images\forward.png", @"Default\Images\forward_pressed.png", true);
                RazerManager.Instance.EnableDynamicKey(RazerAPI.DynamicKeyType.DK9, dkStop, @"Default\Images\stop.png", @"Default\Images\stop_pressed.png", true);

                RazerManager.Instance.EnableDynamicKey(RazerAPI.DynamicKeyType.DK1, dkiTunes, @"Default\Images\iTunes.png", @"Default\Images\iTunes.png", true);

                updateVolumeLabel();
                RazerManager.Instance.EnableDynamicKey(RazerAPI.DynamicKeyType.DK5, dkQuieter, @"Default\Images\quiet.png", @"Default\Images\quiet_pressed.png", true);
                RazerManager.Instance.EnableDynamicKey(RazerAPI.DynamicKeyType.DK10, dkLouder, @"Default\Images\louder.png", @"Default\Images\louder_pressed.png", true);

                //RazerManager.Instance.EnableDynamicKey(RazerAPI.DynamicKeyType.DK4, dkQuieter, @"Default\Images\quiet.png", @"Default\Images\quiet_pressed.png", true);
                //RazerManager.Instance.EnableDynamicKey(RazerAPI.DynamicKeyType.DK5, dkLouder, @"Default\Images\louder.png", @"Default\Images\louder_pressed.png", true);

                //setMuteButton(true);

                setup();
                displayTrack();

                _manager.Touchpad.SetWindow(this, Touchpad.RenderMethod.Polling);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception: " + ex.Message, "Error!", MessageBoxButton.OK, MessageBoxImage.Error);
                releaseiTunes();
                closeApp();
            }
        }

        void _itunes_OnQuittingEvent()
        {
            releaseiTunes();
        }

        void _itunes_OnSoundVolumeChangedEvent(int newVolume)
        {
            updateVolumeLabel();
            _volume = newVolume;
            //setMuteButton();
        }

        private void updateVolumeLabel()
        {
            if (lblVolume.Dispatcher.CheckAccess())
            {
                lblVolume.Content = "Volume: " + _itunes.SoundVolume;
            }
            else
            {
                Dispatcher.BeginInvoke(new Action(() => { displayTrack(); }));
            }


        }

        private void dkiTunes(object sender, EventArgs e)
        {
            createiTunes();

        }

        private void createiTunes(bool init = false)
        {
            if (_itunes == null)
            {
                _itunes = new iTunesLib.iTunesApp();
                _itunes.OnPlayerPlayingTrackChangedEvent += new _IiTunesEvents_OnPlayerPlayingTrackChangedEventEventHandler(OnPlayerPlayingTrackChangedEvent);
                _itunes.OnPlayerPlayEvent += new _IiTunesEvents_OnPlayerPlayEventEventHandler(OnPlayerPlayEvent);
                _itunes.OnAboutToPromptUserToQuitEvent += new _IiTunesEvents_OnAboutToPromptUserToQuitEventEventHandler(OnAboutToPromptUserToQuitEvent);
                _itunes.OnQuittingEvent += _itunes_OnQuittingEvent;
                _itunes.OnSoundVolumeChangedEvent += _itunes_OnSoundVolumeChangedEvent;

                showHide(false);

                if (!init)
                {
                    displayTrack();
                    setPlayPauseButton();
                }
            }
        }

        private void dkNull(object sender, EventArgs e)
        {

        }


        private void dkBack(object sender, EventArgs e)
        {
            _itunes.BackTrack();
            displayTrack();
        }

        private void dkLouder(object sender, EventArgs e)
        {
            int vol = Math.Min(_itunes.SoundVolume + 5, 100);
            _itunes.SoundVolume = vol;
            _itunes_OnSoundVolumeChangedEvent(vol);
            //setMuteButton();
        }

        private void dkQuieter(object sender, EventArgs e)
        {
            int vol = Math.Max(_itunes.SoundVolume - 5, 0);
            _itunes.SoundVolume = vol;
            _itunes_OnSoundVolumeChangedEvent(vol);
            //setMuteButton();
        }

        private void dkNext(object sender, EventArgs e)
        {
            _itunes.NextTrack();
            displayTrack();
        }
        private void dkPlayPause(object sender, EventArgs e)
        {
            _itunes.PlayPause();
            setPlayPauseButton();
            displayTrack();
        }
        private void dkStop(object sender, EventArgs e)
        {
            _itunes.Stop();
            setPlayPauseButton();
            displayTrack();
        }

        //private void setMuteButton(bool init = false)
        //{
        //    if (init)
        //    {
        //        _muted = (_itunes.SoundVolume == 0);
        //        if (_itunes.SoundVolume == 0)
        //        {
        //            RazerManager.Instance.EnableDynamicKey(RazerAPI.DynamicKeyType.DK10, dkMute, @"Default\Images\sound.png", @"Default\Images\sound_pressed.png", true);
        //        }
        //        else
        //        {
        //            RazerManager.Instance.EnableDynamicKey(RazerAPI.DynamicKeyType.DK10, dkMute, @"Default\Images\mute.png", @"Default\Images\mute_pressed.png", true);
        //        }
        //    }
        //    else
        //    {
        //        if ((_muted && (_itunes.SoundVolume > 0)) || (!_muted && (_itunes.SoundVolume == 0)))
        //        {
        //            _muted = (_itunes.SoundVolume == 0);
        //            if (_itunes.SoundVolume == 0)
        //            {
        //                RazerManager.Instance.EnableDynamicKey(RazerAPI.DynamicKeyType.DK10, dkMute, @"Default\Images\sound.png", @"Default\Images\sound_pressed.png", true);
        //            }
        //            else
        //            {
        //                RazerManager.Instance.EnableDynamicKey(RazerAPI.DynamicKeyType.DK10, dkMute, @"Default\Images\mute.png", @"Default\Images\mute_pressed.png", true);
        //            }
        //        }
        //    }


        //    if (!_muted && (_itunes.SoundVolume == 0))
        //    {
        //        RazerManager.Instance.EnableDynamicKey(RazerAPI.DynamicKeyType.DK10, dkMute, @"Default\Images\sound.png", @"Default\Images\sound_pressed.png", true);
        //    }
        //    else if (_muted && (_itunes.SoundVolume > 0))
        //    {
        //        RazerManager.Instance.EnableDynamicKey(RazerAPI.DynamicKeyType.DK10, dkMute, @"Default\Images\mute.png", @"Default\Images\mute_pressed.png", true);
        //    }
        //    _muted = (_itunes.SoundVolume == 0);
        //}

        private void setPlayPauseButton()
        {
            if (_itunes.PlayerState == ITPlayerState.ITPlayerStatePlaying)
            {
                RazerManager.Instance.EnableDynamicKey(RazerAPI.DynamicKeyType.DK7, dkPlayPause, @"Default\Images\paused.png", @"Default\Images\paused_pressed.png", true);
            }
            else
            {
                RazerManager.Instance.EnableDynamicKey(RazerAPI.DynamicKeyType.DK7, dkPlayPause, @"Default\Images\play.png", @"Default\Images\play_pressed.png", true);
            }
        }

        private void dkMute(object sender, EventArgs e)
        {
            if (_itunes.SoundVolume == 0)
            {
                if (_volume == 0)
                {
                    _volume = 50;
                }
                _itunes.SoundVolume = _volume;
            }
            else
            {
                _volume = _itunes.SoundVolume;
                _itunes.SoundVolume = 0;
            }
            //setMuteButton();
            updateVolumeLabel();
        }

        string addTrailingSlash(string str)
        {
            return str.TrimEnd(System.IO.Path.DirectorySeparatorChar) + System.IO.Path.DirectorySeparatorChar;
        }

        void setup()
        {
            string appDataPath = addTrailingSlash(Environment.GetEnvironmentVariable("LocalAppData"));
            string authorPath = addTrailingSlash(appDataPath + _AUTHOR);
            string brandPath = addTrailingSlash(authorPath + _BRAND);
            string appPath = addTrailingSlash(brandPath + _APPNAME);
            if (!Directory.Exists(authorPath))
            {
                Directory.CreateDirectory(authorPath);
                Directory.CreateDirectory(brandPath);
                Directory.CreateDirectory(appPath);
            }

            if (!Directory.Exists(brandPath))
            {
                Directory.CreateDirectory(brandPath);
                Directory.CreateDirectory(appPath);
            }

            if (!Directory.Exists(appPath))
            {
                Directory.CreateDirectory(appPath);
            }

            DirectoryInfo info = new DirectoryInfo(appPath);
            foreach (FileInfo file in info.GetFiles())
            {
                file.Delete();
            }

            _appPath = appPath;
        }

        void OnAboutToPromptUserToQuitEvent()
        {
            releaseiTunes();
        }

        void OnPlayerPlayEvent(object iTrack)
        {
            displayTrack();
        }

        void OnPlayerPlayingTrackChangedEvent(object iTrack)
        {
            displayTrack();
        }

        void displayTrack()
        {
            if (lblCurrentSong.Dispatcher.CheckAccess() && lblArtist.Dispatcher.CheckAccess())
            {
                IITTrack currentTrack = _itunes.CurrentTrack;
                if (currentTrack != null)
                {
                    if ((_itunes.PlayerState == ITPlayerState.ITPlayerStatePlaying) ||
                        (_itunes.PlayerState == ITPlayerState.ITPlayerStateStopped))
                    {
                        lblCurrentSong.Content = currentTrack.Name;
                        lblArtist.Content = currentTrack.Artist;

                        setArtWork();
                    }
                }
                else
                {
                    lblCurrentSong.Content = "";
                    lblArtist.Content = "";

                    setArtWork();
                }
            }
            else
            {
                Dispatcher.BeginInvoke(new Action(() => { displayTrack(); }));
            }
        }

        void setArtWork()
        {
            if (imgArtwork.Dispatcher.CheckAccess())
            {
                //already checked the player is playing
                IITTrack currentTrack = _itunes.CurrentTrack;
                if (currentTrack != null)
                {
                    IITArtworkCollection _colArtwork = currentTrack.Artwork;
                    Uri source;
                    if (_colArtwork.Count >= 1)
                    {
                        IITArtwork _artwork = _colArtwork[1];
                        string filename = _appPath + currentTrack.Artist + "-" + currentTrack.Name + ".jpg";
                        try
                        {
                            _artwork.SaveArtworkToFile(filename);
                        }
                        catch (Exception ex)
                        {
                        }
                        source = new Uri(filename, UriKind.Absolute);
                    }
                    else
                    {
                        //No Artwork
                        source = new Uri("Default\\Images\\noArtwork.jpg", UriKind.Relative);
                    }

                    BitmapImage bi = new BitmapImage();
                    bi.BeginInit();
                    bi.UriSource = source;
                    bi.EndInit();
                    imgArtwork.Stretch = Stretch.Fill;
                    imgArtwork.Source = bi;
                    imgArtwork.Height = 300;
                    imgArtwork.Width = 300;

                }
                else
                {
                    Uri source = new Uri(@"Default\Images\noArtwork.jpg", UriKind.Relative);
                    BitmapImage bi = new BitmapImage();
                    bi.BeginInit();
                    bi.UriSource = source;
                    bi.EndInit();
                    imgArtwork.Stretch = Stretch.Fill;
                    imgArtwork.Source = bi;
                    imgArtwork.Height = 300;
                    imgArtwork.Width = 300;
                }
            }
            else
            {
                Dispatcher.BeginInvoke(new Action(() => { setArtWork(); }));
            }
        }

        void OnAppEvent(object sender, SharpBlade.Razer.Events.AppEventEventArgs e)
        {
            if (
                (e.Type == SharpBlade.Native.RazerAPI.AppEventType.Deactivated) ||
                (e.Type == SharpBlade.Native.RazerAPI.AppEventType.Exit) ||
                (e.Type == SharpBlade.Native.RazerAPI.AppEventType.Close))
            {
                releaseiTunes();
                closeApp();
            }
        }

        void releaseiTunes()
        {
            _itunes.OnPlayerPlayingTrackChangedEvent -= OnPlayerPlayingTrackChangedEvent;
            _itunes.OnPlayerPlayEvent -= OnPlayerPlayEvent;
            _itunes.OnAboutToPromptUserToQuitEvent -= OnAboutToPromptUserToQuitEvent;
            _itunes.OnQuittingEvent -= _itunes_OnQuittingEvent;
            _itunes.OnSoundVolumeChangedEvent -= _itunes_OnSoundVolumeChangedEvent;
            Marshal.ReleaseComObject(_itunes);
            _itunes = null;
            showHide();
        }

        private void showHide(bool hidden = true)
        {
            if (lblVolume.Dispatcher.CheckAccess() && 
                lblArtist.Dispatcher.CheckAccess() && 
                lblCurrentSong.Dispatcher.CheckAccess() && 
                imgArtwork.Dispatcher.CheckAccess())
            {
                if (hidden)
                {
                    lblVolume.Visibility = System.Windows.Visibility.Hidden;
                    lblArtist.Visibility = System.Windows.Visibility.Hidden;
                    lblCurrentSong.Visibility = System.Windows.Visibility.Hidden;
                    imgArtwork.Visibility = System.Windows.Visibility.Hidden;

                    lblClosed.Visibility = System.Windows.Visibility.Visible;
                    lblAction.Visibility = System.Windows.Visibility.Visible;
                    imgiTunes.Visibility = System.Windows.Visibility.Visible;
                }
                else
                {
                    lblVolume.Visibility = System.Windows.Visibility.Visible;
                    lblArtist.Visibility = System.Windows.Visibility.Visible;
                    lblCurrentSong.Visibility = System.Windows.Visibility.Visible;
                    imgArtwork.Visibility = System.Windows.Visibility.Visible;

                    lblClosed.Visibility = System.Windows.Visibility.Hidden;
                    lblAction.Visibility = System.Windows.Visibility.Hidden;
                    imgiTunes.Visibility = System.Windows.Visibility.Hidden;
                }
            }
            else
            {
                Dispatcher.BeginInvoke(new Action(() => { showHide(hidden); }));
            }
        }

        void closeApp()
        {
            Application.Current.Shutdown();
        }
    }
}
