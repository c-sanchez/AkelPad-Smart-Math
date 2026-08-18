#ifndef __PTRARRAY_BI__
#define __PTRARRAY_BI__

const PTR_ARRAY_CHUNK as Integer = 16

type PtrDeleterType as sub(byval as Any Ptr)
type PtrMatcherType as function(byval as Any Ptr, byval as Any Ptr) as Boolean

type PtrArray
public:
  declare constructor()
  declare constructor(byval customDeleter as PtrDeleterType, byval customMatcher as PtrMatcherType)
  declare destructor()
  declare operator [](byval index as Integer) as Any Ptr
  declare sub Clear()
  declare sub Append(byval p as Any Ptr)
  declare function Find(byval p as Any Ptr) as Integer
  declare sub RemoveAt(byval index as Integer)
  declare function Count() as Integer
  declare function Capacity() as Integer
  declare function Items() as Any Ptr Ptr

private:
  declare constructor(byref rhs as PtrArray)
  declare operator let(byref rhs as PtrArray)

  m_items(any) as Any Ptr
  m_count as Integer
  m_capacity as Integer
  m_deleter as PtrDeleterType
  m_matcher as PtrMatcherType
end type

constructor PtrArray()
  m_count = 0
  m_capacity = PTR_ARRAY_CHUNK
  m_deleter = 0
  m_matcher = 0
  redim m_items(0 to m_capacity - 1)
end constructor

constructor PtrArray(byval customDeleter as PtrDeleterType, byval customMatcher as PtrMatcherType)
  m_count = 0
  m_capacity = PTR_ARRAY_CHUNK
  m_deleter = customDeleter
  m_matcher = customMatcher
  redim m_items(0 to m_capacity - 1)
end constructor

destructor PtrArray()
  Clear()
  erase m_items
  m_capacity = 0
end destructor

operator PtrArray.[](byval index as Integer) as Any Ptr
  if index < 0 orelse index >= m_count then return 0
  return m_items(index)
end operator

sub PtrArray.Clear()
  if m_deleter <> 0 then
    for i as Integer = 0 to m_count - 1
      m_deleter(m_items(i))
      m_items(i) = 0
    next
  end if
  m_count = 0
end sub

sub PtrArray.Append(byval p as Any Ptr)
  if m_capacity = 0 then
    m_capacity = PTR_ARRAY_CHUNK
    redim m_items(0 to m_capacity - 1)
  elseif m_count >= m_capacity then
    m_capacity += PTR_ARRAY_CHUNK
    redim preserve m_items(0 to m_capacity - 1)
  end if

  m_items(m_count) = p
  m_count += 1
end sub

function PtrArray.Find(byval p as Any Ptr) as Integer
  for i as Integer = 0 to m_count - 1
    if m_matcher <> 0 then
      if m_matcher(m_items(i), p) then return i
    else
      if m_items(i) = p then return i
    end if
  next

  return -1 ' not found
end function

sub PtrArray.RemoveAt(byval index as Integer)
  if index < 0 orelse index >= m_count then exit sub

  dim pRemoved as Any Ptr = m_items(index)

  for i as Integer = index to m_count - 2
    m_items(i) = m_items(i + 1)
  next

  m_count -= 1
  m_items(m_count) = 0

  if m_deleter <> 0 then m_deleter(pRemoved)
end sub

function PtrArray.Count() as Integer
  return m_count
end function

function PtrArray.Capacity() as Integer
  return m_capacity
end function

function PtrArray.Items() as Any Ptr Ptr
  if m_capacity <= 0 then return 0
  return @m_items(0)
end function

#endif
