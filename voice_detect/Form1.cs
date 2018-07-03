using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

using SpeechLib;
using iTunesLib;

namespace voice_detect
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //VoiceDetect
        private SpeechLib.SpInProcRecoContext mRecognizerRule = null;
        private SpeechLib.ISpeechRecoGrammar mRecognizerGrammarRule = null;
        private SpeechLib.ISpeechGrammarRule mRecognizerGrammarRuleGrammarRule = null;

        //iTunes
        private iTunesApp itunes = null;

        private void initRecognizer()
        {
            mRecognizerRule = new SpeechLib.SpInProcRecoContext();
            bool hit = false;
            foreach (SpObjectToken recoPerson in mRecognizerRule.Recognizer.GetRecognizers())
            {
                if (recoPerson.GetAttribute("Language") == "411")
                {
                    mRecognizerRule.Recognizer.Recognizer = recoPerson;
                    hit = true;
                    break;
                }
            }
            if (!hit)
            {
                MessageBox.Show("日本語認識が利用できません");
                Application.Exit();
            }

            mRecognizerRule.Recognizer.AudioInput = createMicrofon();

            if (mRecognizerRule.Recognizer.AudioInput == null)
            {
                MessageBox.Show("マイク初期化エラー");
                Application.Exit();
            }

            mRecognizerRule.Hypothesis +=
                delegate(int streamNumber, object streamPosition, SpeechLib.ISpeechRecoResult result)
                {
                    string strText = result.PhraseInfo.GetText();
                    textBox1.Text = strText;
                };
            mRecognizerRule.Recognition +=
                delegate(int streamNumber, object streamPosition, SpeechLib.SpeechRecognitionType srt, SpeechLib.ISpeechRecoResult isrr)
                {
                    SpeechEngineConfidence confidence = isrr.PhraseInfo.Rule.Confidence;
                    switch (confidence)
                    {
                        case SpeechEngineConfidence.SECHighConfidence:
                            label3.Text = "Confidence is High";
                            break;
                        case SpeechEngineConfidence.SECNormalConfidence:
                            label3.Text = "Confidence is Normal";
                            break;
                        case SpeechEngineConfidence.SECLowConfidence:
                            label3.Text = "Confidence is Low";
                            textBox2.Text = "信頼性が低すぎます";
                            return;
                    }
                    string strText = isrr.PhraseInfo.GetText();
                    //isrr.PhraseInfo.
                    label4.Text = isrr.RecoContext.Voice.Volume.ToString();
                    if (strText == "えんいー")
                    {
                        Application.Exit();
                    }

                    if (itunes != null)
                    {
                        switch (strText)
                        {
                            case "あいちゅーんず．つぎのきょく":
                            case "あいちゅーんず．つぎ":
                                itunes.NextTrack();
                                break;
                            case "あいちゅーんず．まえのきょく":
                            case "あいちゅーんず．まえ":
                                itunes.PreviousTrack();
                                break;
                            case "あいちゅーんず．いちじていし":
                                itunes.Pause();
                                break;
                            case "あいちゅーんず．ていし":
                                itunes.Stop();
                                break;
                            case "あいちゅーんず．さいせい":
                                itunes.Play();
                                break;
                            case "あいちゅーんず．しね":
                                itunes.Quit();
                                unhockiTunes();
                                break;
                            case "あいちゅーんず．しずかに":
                                itunes.SoundVolume = 50;
                                break;
                            case "あいちゅーんず．おおきく":
                                itunes.SoundVolume = 100;
                                break;
                            case "あいちゅーんず．らんだむ":
                                itunes.CurrentPlaylist.Shuffle = !itunes.CurrentPlaylist.Shuffle;
                                break;
                        }
                    }
                    else
                    {
                        if (strText == "あいちゅーんず．おきろ")
                        {
                            initiTunes();
                        }
                    }
                    textBox2.Text = strText;
                };
            mRecognizerRule.StartStream +=
                delegate(int streamNumber, object streamPosition)
                {
                    textBox1.Text = textBox2.Text = "";
                };
            mRecognizerRule.FalseRecognition +=
                delegate(int streamNumber, object streamPosition, SpeechLib.ISpeechRecoResult isrr)
                {
                    textBox1.Text = textBox2.Text = label3.Text = "--Error!--";
                };

            mRecognizerGrammarRule = mRecognizerRule.CreateGrammar();
            mRecognizerGrammarRule.Reset();
            mRecognizerGrammarRuleGrammarRule = mRecognizerGrammarRule.Rules.Add("TopLevelRule", SpeechRuleAttributes.SRATopLevel | SpeechRuleAttributes.SRADynamic);

            mRecognizerGrammarRuleGrammarRule.InitialState.AddWordTransition(null, "あいちゅーんず．おきろ");
            mRecognizerGrammarRuleGrammarRule.InitialState.AddWordTransition(null, "あいちゅーんず．つぎのきょく");
            mRecognizerGrammarRuleGrammarRule.InitialState.AddWordTransition(null, "あいちゅーんず．まえのきょく");
            mRecognizerGrammarRuleGrammarRule.InitialState.AddWordTransition(null, "あいちゅーんず．つぎ");
            mRecognizerGrammarRuleGrammarRule.InitialState.AddWordTransition(null, "あいちゅーんず．まえ");
            mRecognizerGrammarRuleGrammarRule.InitialState.AddWordTransition(null, "あいちゅーんず．いちじていし");
            mRecognizerGrammarRuleGrammarRule.InitialState.AddWordTransition(null, "あいちゅーんず．ていし");
            mRecognizerGrammarRuleGrammarRule.InitialState.AddWordTransition(null, "あいちゅーんず．さいせい");
            mRecognizerGrammarRuleGrammarRule.InitialState.AddWordTransition(null, "あいちゅーんず．しね");
            mRecognizerGrammarRuleGrammarRule.InitialState.AddWordTransition(null, "あいちゅーんず．しずかに");
            mRecognizerGrammarRuleGrammarRule.InitialState.AddWordTransition(null, "あいちゅーんず．おおきく");
            mRecognizerGrammarRuleGrammarRule.InitialState.AddWordTransition(null, "あいちゅーんず．らんだむ");
           // mRecognizerGrammarRuleGrammarRule.InitialState.AddWordTransition(null, "えんいー");

            mRecognizerGrammarRule.Rules.Commit();
            mRecognizerGrammarRule.CmdSetRuleState("TopLevelRule", SpeechRuleState.SGDSActive);
        }

        private SpeechLib.SpObjectToken createMicrofon()
        {
            var ObjectTokenCat = new SpeechLib.SpObjectTokenCategory();
            ObjectTokenCat.SetId(@"HkEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\AudioInput");
            var token = new SpeechLib.SpObjectToken();
            token.SetId(ObjectTokenCat.Default);

            return token;
            //SpeechLib.SpObjectTokenCategory objAudioTokenCategory = new SpeechLib.SpObjectTokenCategory();
            //objAudioTokenCategory.SetId(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech Server\v11.0\AudioInput", false);
            //SpeechLib.SpObjectToken objAudioToken = new SpeechLib.SpObjectToken();
            //objAudioToken.SetId(objAudioTokenCategory.Default, @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech Server\v11.0\AudioInput", false);
            //return objAudioToken;
        }

        //再生イベント
        void itunes_OnPlayerPlayEvent(object iTrack)
        {
        }

        //終了検知
        void itunes_OnAboutToPromptUserToQuitEvent()
        {
            unhockiTunes();
            this.Invoke((System.Windows.Forms.MethodInvoker)delegate()
            {
                this.Close();
            });
        }

        //初期化
        private void initiTunes()
        {
            itunes = new iTunesApp();
            itunes.OnAboutToPromptUserToQuitEvent +=
                new _IiTunesEvents_OnAboutToPromptUserToQuitEventEventHandler(itunes_OnAboutToPromptUserToQuitEvent);
            itunes.OnPlayerPlayEvent +=
                new _IiTunesEvents_OnPlayerPlayEventEventHandler(itunes_OnPlayerPlayEvent);

        }

        //ホック解除
        private void unhockiTunes()
        {
            itunes.OnPlayerPlayEvent -=
                new _IiTunesEvents_OnPlayerPlayEventEventHandler(itunes_OnPlayerPlayEvent);
            itunes.OnAboutToPromptUserToQuitEvent -=
                new _IiTunesEvents_OnAboutToPromptUserToQuitEventEventHandler(itunes_OnAboutToPromptUserToQuitEvent);

            Marshal.ReleaseComObject(itunes);
            itunes = null;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            initRecognizer();
        }

        ~Form1()
        {
            if (itunes != null)
            {
                unhockiTunes();
            }
        }
    }
}
