
using Documenter, XLSX

makedocs(
    format = :html,
    sitename = "XLSX.jl",
    modules = [ XLSX ],
    pages = [ "index.md", "api.md" ]
)

deploydocs(
    repo = "github.com/felipenoris/XLSX.jl.git",
    target = "build",
    julia  = "1.0",
    deps   = nothing,
    make   = nothing
)
