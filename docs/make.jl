
using Documenter, XLSX

makedocs(
    format = :html,
    sitename = "XLSX.jl",
    modules = [XLSX],
    pages = ["index.md",
            "api.md"
             ]
)

deploydocs(
    repo = "github.com/felipenoris/XLSX.jl.git",
    target = "build",
    julia  = "0.6",
    deps   = nothing,
    make   = nothing
)
