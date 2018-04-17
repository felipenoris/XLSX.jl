
function _cellvalue_datetime(v::AbstractString, _is_date_1904::Bool) :: Union{Dates.DateTime, Dates.Date, Dates.Time}
    if contains(v, ".")
        time_value = parse(Float64, v)
        @assert time_value >= 0

        if time_value <= 1
            # Time
            return _time(time_value)
        else
            # DateTime
            return _datetime(time_value, _is_date_1904)
        end
    else
        # Date
        return _date(parse(Int, v), _is_date_1904)
    end
end

"""
Converts Excel number to Time.
`x` must be between 0 and 1.

To represent Time, Excel uses the decimal part
of a floating point number. `1` equals one day.
"""
function _time(x::Float64) :: Dates.Time
    @assert x >= 0 && x <= 1
    return Dates.Time(Dates.Nanosecond(round(Int, x * 86400) * 1E9 ))
end

"""
Converts Excel number to Date.

See also: `isdate1904` function.
"""
function _date(x::Int, _is_date_1904::Bool) :: Dates.Date
    if _is_date_1904
        return Date(Dates.rata2datetime(x + 695056))
    else
        return Date(Dates.rata2datetime(x + 693594))
    end
end

"""
Converts Excel number to DateTime.

The decimal part represents the Time (see `_time` function).
The integer part represents the Date.

See also: `isdate1904` function.
"""
function _datetime(x::Float64, _is_date_1904::Bool) :: Dates.DateTime
    @assert x >= 0

    local dt::Dates.Date
    local hr::Dates.Time

    dt_part = trunc(Int, x)
    hr_part = x - dt_part

    dt = _date(dt_part, _is_date_1904)
    hr = _time(hr_part)

    return dt + hr
end
