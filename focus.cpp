#define NAPI_CPP_EXCEPTIONS 1

#include <napi.h>
#include <winrt/Windows.Foundation.h>
#include <winrt/Windows.UI.Shell.h>

using namespace winrt;
using namespace Windows::UI::Shell;

class FocusAssist : public Napi::ObjectWrap<FocusAssist> {
    public:
        static Napi::Object Init(Napi::Env env, Napi::Object exports);
        FocusAssist(const Napi::CallbackInfo& info);
        Napi::Value IsEnabled(const Napi::CallbackInfo& info);
    private:
        static Napi::FunctionReference constructor;
        FocusSessionManager manager;
};

Napi::FunctionReference FocusAssist::constructor;

FocusAssist::FocusAssist(const Napi::CallbackInfo& info) : Napi::ObjectWrap<FocusAssist>(info), manager(FocusSessionManager::GetDefault()) {
}

Napi::Value FocusAssist::IsEnabled(const Napi::CallbackInfo& info){
    Napi::Env env = info.Env();

    return Napi::Boolean::New(env, manager.IsFocusActive());
}

Napi::Object FocusAssist::Init(Napi::Env env, Napi::Object exports){
    Napi::HandleScope scope(env);

    Napi::Function func = DefineClass(env, "FocusAssist", {
        InstanceMethod("isEnabled", &FocusAssist::IsEnabled)
    });

    constructor = Napi::Persistent(func);
    constructor.SuppressDestruct();

    exports.Set("FocusAssist", func);
    return exports;
}

Napi::Object Init(Napi::Env env, Napi::Object exports){
    winrt::init_apartment(winrt::apartment_type::multi_threaded);
    FocusAssist::Init(env, exports);
    return exports;
}

NODE_API_MODULE(NODE_GYP_MODULE_NAME, Init);