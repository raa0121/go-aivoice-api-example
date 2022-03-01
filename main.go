package main

import (
	"bufio"
	"encoding/json"
	"fmt"
	"log"
	"math/rand"
	"os"
	"time"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

const (
	HostStatus_NotRunning = iota
	HostStatus_NotConnected
	HostStatus_Idle
	HostStatus_Busy
)

type MasterControl struct {
	Volume            float64 `json:"Volume"`
	Speed             float64 `json:"Speed"`
	Pitch             float64 `json:"Pitch"`
	PitchRange        float64 `json:"PitchRange"`
	IsPauseEnabled    bool    `json:"IsPauseEnabled"`
	MiddlePause       int     `json:"MiddlePause"`
	LongPause         int     `json:"LongPause"`
	SentencePause     int     `json:"SentencePause"`
	VolumeDecibel     float64 `json:"VolumeDecibel"`
	PitchCent         float64 `json:"PitchCent"`
	PitchHalfTone     float64 `json:"PitchHalfTone"`
	PitchRangePercent int     `json:"PitchRangePercent"`
}

var StatusMap = map[int64]string{
	HostStatus_NotConnected: "NotConnected",
	HostStatus_Idle: "Idle",
	HostStatus_Busy: "Busy",
}

func main() {
	ret := 0

	err := ole.CoInitialize(0)
	if err != nil && FAILED(err) {
		fmt.Println(err)
		fmt.Println("Press Enter")
		input := bufio.NewScanner(os.Stdin)
		input.Scan()
		os.Exit(-1)
	}

	unknown, err := oleutil.CreateObject("AI.Talk.Editor.Api.TtsControl")
	if err != nil && FAILED(err) {
		fmt.Println(err)
		ret = -2
	}

	aivoice, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil && FAILED(err) {
		fmt.Println(err)
		ret = -2
	}

	hosts, err := oleutil.CallMethod(aivoice, "GetAvailableHostNames")
	if err != nil && FAILED(err) {
		fmt.Println("ITtsControl::Initialize() Failed.")
		ret = -3
	}
	if len(hosts.ToArray().ToValueArray()) < 0 {
		fmt.Println("Host not found.")
		ret = -3
	}

	_, err = oleutil.CallMethod(aivoice, "Initialize", hosts.ToArray().ToValueArray()[0])
	if err != nil && FAILED(err) {
		fmt.Println("ITtsControl::Initialize() Failed.")
		ret = -3
	}

	initialized, err := oleutil.GetProperty(aivoice, "IsInitialized")
	fmt.Printf("ITtsControl::IsInitialized : %t\n", initialized.Value())

	if false == initialized.Value() {
		fmt.Println("ITtsControl::IsInitialized = false.")
		ret = -4
	} else {
		fmt.Println("ITtsControl::Initialize() Succeeded.")
	}

	status, err := oleutil.GetProperty(aivoice, "Status")
	if status.Val == HostStatus_NotRunning {
		_, err = oleutil.CallMethod(aivoice, "StartHost")
		if err != nil && FAILED(err) {
			fmt.Printf("ITtsControl::StartHost() Failed.\n")
			ret = -5
		} else {
			fmt.Printf("ITtsControl::StartHost() Succeeded.\n")
		}
	}

	status = oleutil.MustGetProperty(aivoice, "Status")
	if status.Val == HostStatus_NotConnected {
		_, err = oleutil.CallMethod(aivoice, "Connect")
		if err != nil && FAILED(err) {
			fmt.Printf("ITtsControl::Connect() Failed.\n")
			ret = -6
		} else {
			fmt.Printf("ITtsControl::Connect() Succeeded.\n")
		}
	}

	if ret < 0 {
		fmt.Printf("Sample program end : %d\n", ret)
		fmt.Println("Press Enter")
		input := bufio.NewScanner(os.Stdin)
		input.Scan()
		os.Exit(-1)
	}

	version, err := oleutil.GetProperty(aivoice, "Version")
	fmt.Println("ITtsControl::Version : " + version.ToString())

	voices, err := oleutil.GetProperty(aivoice, "VoiceNames")
	fmt.Println("ITtsControl::VoiceNames : ")
	for _, voice := range voices.ToArray().ToValueArray() {
		fmt.Println(voice)
	}

	presets, err := oleutil.GetProperty(aivoice, "VoicePresetNames")
	fmt.Println("ITtsControl::VoicePresetNames :")
	for _, preset := range presets.ToArray().ToValueArray() {
		fmt.Println(preset)
	}

	status = oleutil.MustGetProperty(aivoice, "Status")
	fmt.Println("ITtsControl::Status : " + StatusMap[status.Val])
	currentVoicePresetName, err := oleutil.GetProperty(aivoice, "CurrentVoicePresetName")
	fmt.Println("ITtsControl::CurrentVoicePresetName: " + currentVoicePresetName.ToString())

	text, err := oleutil.GetProperty(aivoice, "Text")
	fmt.Println("ITtsControl::Text : " + text.ToString())

	textSelectionStart, err := oleutil.GetProperty(aivoice, "TextSelectionStart")
	fmt.Printf("ITtsControl::textSelectionStart : %d\n", textSelectionStart.Value())

	textSelectionLength, err := oleutil.GetProperty(aivoice, "TextSelectionLength")
	fmt.Printf("ITtsControl::textSelectionLength : %d\n", textSelectionLength.Value())

	mc, err := oleutil.GetProperty(aivoice, "MasterControl")
	fmt.Printf("ITtsControl::MasterControl : %s\n", mc.Value())

	oleutil.MustPutProperty(aivoice, "CurrentVoicePresetName", currentVoicePresetName.ToString())
	fmt.Println("ITtsControl::CurrentVoicePresetName = \"" + oleutil.MustGetProperty(aivoice, "CurrentVoicePresetName").ToString() + "\"")

	oleutil.MustPutProperty(aivoice, "Text", "メロスは激怒した。必ず、かの邪智暴虐の王を除かなければならぬと決意した。メロスには政治がわからぬ。メロスは、村の牧人である。笛を吹き、羊と遊んで暮して来た。けれども邪悪に対しては、人一倍に敏感であった。")
	fmt.Println("ITtsControl::Text= \"" + oleutil.MustGetProperty(aivoice, "Text").ToString() + "\"")

	oleutil.MustPutProperty(aivoice, "TextSelectionStart", textSelectionStart.Value())
	fmt.Printf("ITtsControl::TextSelectionStart = %d\n", oleutil.MustGetProperty(aivoice, "TextSelectionStart").Value())

	oleutil.MustPutProperty(aivoice, "TextSelectionLength", textSelectionLength.Value())
	fmt.Printf("ITtsControl::TextSelectionLength = %d\n", oleutil.MustGetProperty(aivoice, "TextSelectionLength").Value())

	rand.Seed(time.Now().UnixNano())
	master := MasterControl{}
	if rand.Intn(2) == 0 {
		master.Volume = 1.2
	} else {
		master.Volume = 0.8
	}
	if rand.Intn(2) == 0 {
		master.Speed = 1.2
	} else {
		master.Speed = 0.8
	}
	if rand.Intn(2) == 0 {
		master.Pitch = 1.2
	} else {
		master.Pitch = 0.8
	}
	if rand.Intn(2) == 0 {
		master.PitchRange = 1.2
	} else {
		master.PitchRange = 0.8
	}
	if rand.Intn(2) == 0 {
		master.MiddlePause = int(150.0 * 1.2)
	} else {
		master.MiddlePause = int(150.0 * 0.8)
	}
	if rand.Intn(2) == 0 {
		master.LongPause = int(370.0 * 1.2)
	} else {
		master.LongPause = int(370.0 * 0.8)
	}
	if rand.Intn(2) == 0 {
		master.SentencePause = int(800.0 * 1.2)
	} else {
		master.LongPause = int(800.0 * 0.8)
	}

	j, err := json.Marshal(master)
	if err != nil {
		log.Fatal(err)
	}
	oleutil.MustPutProperty(aivoice, "MasterControl", string(j))
	fmt.Println("ITtsControl::MasterControl = " + oleutil.MustGetProperty(aivoice, "MasterControl").ToString())

	_, err = oleutil.CallMethod(aivoice, "Play")
	if err != nil && FAILED(err) {
		fmt.Println("ITtsControl::Play() Failed.")
		ret = -5
	} else {
		fmt.Println("ITtsControl::Play() Succeeded.")
	}

	time.Sleep(10000 * time.Millisecond)

	_, err = oleutil.CallMethod(aivoice, "Stop")
	if err != nil && FAILED(err) {
		fmt.Println("ITtsControl::Stop() Failed.")
	} else {
		fmt.Println("ITtsControl::Stop() Succeeded.")
	}

	time.Sleep(1000 * time.Millisecond)

	path := os.Getenv("USERPROFILE") + "/Desktop/test.wav"
	_, err = oleutil.CallMethod(aivoice, "SaveAudioToFile", path)
	if err != nil && FAILED(err) {
		fmt.Println("ITtsControl::SaveAudioToFile(\""+ path +"\") Failet.")
	} else {
		fmt.Println("ITtsControl::SaveAudioToFile(\""+ path +"\") Succeeded.")
	}

	_, err = oleutil.CallMethod(aivoice, "Disconnect")
	if err != nil && FAILED(err) {
	}

	finishPrompt(ret)

	ole.CoUninitialize()
}

func FAILED(err error) bool {
	return err.(*ole.OleError).Code() < 0
}

func finishPrompt(ret int) {
	fmt.Printf("Sample program end. : %d\n", ret)
	fmt.Println("Press Enter")
    input := bufio.NewScanner(os.Stdin)
    input.Scan()
}
