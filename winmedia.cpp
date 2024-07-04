#define NAPI_CPP_EXCEPTIONS 1

#include <napi.h>
#include <winrt/Windows.Foundation.h>
#include <winrt/Windows.Media.Core.h>
#include <winrt/Windows.Media.Playback.h>
#include "util.h"
#include <iostream>

using namespace winrt;
using namespace Windows::Media::Core;
using namespace Windows::Media::Playback;
using namespace Windows::Foundation;

class AudioPlayer : public Napi::ObjectWrap<AudioPlayer> {
    public:
        static Napi::Object Init(Napi::Env env, Napi::Object exports);
        AudioPlayer(const Napi::CallbackInfo& info);
        void Play(const Napi::CallbackInfo& info);
        void Pause(const Napi::CallbackInfo& info);
        void Stop(const Napi::CallbackInfo& info);
        void Finalize(Napi::Env env);

    private:
        static Napi::FunctionReference constructor;
        MediaPlayer player;
        std::string source;
};

Napi::FunctionReference AudioPlayer::constructor;

AudioPlayer::AudioPlayer(const Napi::CallbackInfo& info) : Napi::ObjectWrap<AudioPlayer>(info), player(MediaPlayer()) {
    if( info.Length() < 1 || !info[0].IsString() ){
        Napi::TypeError::New(info.Env(), "Must provide a source").ThrowAsJavaScriptException();
    }

    source = info[0].As<Napi::String>().Utf8Value();
}

void AudioPlayer::Play(const Napi::CallbackInfo& info){
    if(
        player.PlaybackSession().PlaybackState() != MediaPlaybackState::Playing &&
        player.PlaybackSession().PlaybackState() != MediaPlaybackState::Paused
    ) {
        player.Source(
            MediaSource::CreateFromUri(Uri(NarrowStringToWideString(source)))
        );
    }
    
    player.Play();
}

void AudioPlayer::Pause(const Napi::CallbackInfo& info){
    player.Pause();
}

void AudioPlayer::Stop(const Napi::CallbackInfo& info){
    player.Pause();
    player.PlaybackSession().Position(TimeSpan(0));
}

void AudioPlayer::Finalize(Napi::Env env){
    player.Close();
}

Napi::Object AudioPlayer::Init(Napi::Env env, Napi::Object exports){
    Napi::HandleScope scope(env);

    Napi::Function func = DefineClass(env, "AudioPlayer", {
        InstanceMethod("play", &AudioPlayer::Play),
        InstanceMethod("pause", &AudioPlayer::Pause),
        InstanceMethod("stop", &AudioPlayer::Stop)
    });

    constructor = Napi::Persistent(func);
    constructor.SuppressDestruct();

    exports.Set("AudioPlayer", func);
    return exports;
}

Napi::Object Init(Napi::Env env, Napi::Object exports){
    try {
        winrt::init_apartment(winrt::apartment_type::single_threaded);
        AudioPlayer::Init(env, exports);
    } catch (winrt::hresult_error const& ex) {
        Napi::Error::New(env, WideStringToNarrowString(ex.message().c_str())).ThrowAsJavaScriptException();
    }
    return exports;
}

NODE_API_MODULE(NODE_GYP_MODULE_NAME, Init);