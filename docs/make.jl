
using Documenter, XLSX

makedocs(
    sitename = "XLSX.jl",
    modules = [ XLSX ],
    pages = [ "Home" => "index.md",
              "Tutorial" => "tutorial.md",
              "API Reference" => "api.md",
              "Migration Guides" => "migration.md",
            ]
)

deploydocs(
    repo = "github.com/felipenoris/XLSX.jl.git",
    target = "build",
)
