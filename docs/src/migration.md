# Migration Guides

## Migrating Legacy Code to v0.8

Version `v0.8` introduced a breaking change on methods [`XLSX.gettable`](@ref) and [`XLSX.readtable`](@ref).

These methods used to return a tuple `data, column_labels`.
On XLSX `v0.8` these methods return a `XLSX.DataTable` struct that implements [`Tables.jl`](https://github.com/JuliaData/Tables.jl) interface.

### Basic code replacement

Before

```julia
data, col_names = XLSX.readtable(joinpath(data_directory, "general.xlsx"), "table4")
```

After

```julia
dtable = XLSX.readtable(joinpath(data_directory, "general.xlsx"), "table4")
data, col_names = dtable.data, dtable.column_labels
```

### Reading DataFrames

Since `XLSX.DataTable` implements `Tables.jl` interface,
the result of `XLSX.gettable` or `XLSX.readtable` can be
passed to a `DataFrame` constructor.

Before

```julia
df = DataFrame(XLSX.readtable("myfile.xlsx", "mysheet")...)
```

After

```julia
df = DataFrame(XLSX.readtable("myfile.xlsx", "mysheet"))
```
