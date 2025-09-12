
using Documenter, XLSX

makedocs(
    sitename = "XLSX.jl",
    modules = [ XLSX ],
    pages = [
        "Home" => "index.md",
        "Tutorial" => "tutorial.md",
        "Formatting Guide" => Any[
            "Cell formats" => "formatting/cellFormatting.md",
            "Conditional formats" => "formatting/conditionalFormatting.md",
            "Working with merged cells" => "formatting/mergedCells.md",
            "Examples" => "formatting/examples.md"
        ],
        "Migration Guide" => "migration.md",
        "API Reference" => Any[
            "Files and worksheets" => "api/files.md",
            "Cells and data" => "api/data.md",
            "Formats" => "api/formats.md",
        ]
     ],
    checkdocs=:none,
)

deploydocs(
    repo = "github.com/felipenoris/XLSX.jl.git",
    target = "build",
)
