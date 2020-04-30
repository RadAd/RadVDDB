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
