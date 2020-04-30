#include <Windows.h>

#include <string>

#define tstring wstring

class WinError
{
private:
    DWORD           m_Error;
    std::tstring    m_UserMessage;
    HMODULE         m_hModule;

public:
    WinError(HMODULE hModule, const std::tstring& UserMessage = TEXT("Windows Error : "), DWORD Error = GetLastError())
        : m_Error(Error), m_UserMessage(UserMessage), m_hModule(hModule)
    {
    }

    WinError(const std::tstring& UserMessage = TEXT("Windows Error : "), DWORD Error = GetLastError())
        : m_Error(Error), m_UserMessage(UserMessage), m_hModule(NULL)
    {
    }

    DWORD GetError() const
    {
        return m_Error;
    }

    std::tstring GetString() const { return m_UserMessage + GetString(m_Error, m_hModule); }
    static std::tstring GetString(DWORD error, HMODULE hModule = NULL);
};

inline void ThrowWinError(HMODULE hModule, LPCTSTR msg)
{
    throw WinError(hModule, msg);
}
