#pragma once

#include <memory>
#include <assert.h>

template <class T>
class GDIDeleter
{
public:
    typedef T pointer;
    void operator()(T hObject)
    {
        DeleteObject(hObject);
    }
};

template <class T>
class GDIPtr : public std::unique_ptr<T, GDIDeleter<T>>
{
public:
    using std::unique_ptr<T, GDIDeleter<T>>::unique_ptr;

    operator T() const { return std::unique_ptr<T, GDIDeleter<T>>::get(); }
};

inline BOOL Rectangle(HDC hdc, RECT r)
{
    return Rectangle(hdc, r.left, r.top, r.right, r.bottom);
}

inline SIZE GetSize(const RECT r)
{
    return { r.right - r.left, r.bottom - r.top };
}

inline void Multiply(RECT& r, const SIZE sz)
{
    r.left *= sz.cx;
    r.top *= sz.cy;
    r.right *= sz.cx;
    r.bottom *= sz.cy;
}

inline void Divide(RECT& r, const SIZE sz)
{
    r.left /= sz.cx;
    r.top /= sz.cy;
    r.right /= sz.cx;
    r.bottom /= sz.cy;
}

inline double AspectRatio(const SIZE sz)
{
    return (double) sz.cx / sz.cy;
}
