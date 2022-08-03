using ArtifactUtils
using Pkg
using Pkg.PlatformEngines
using Tar

artifact_content = joinpath(@__DIR__() |> dirname, "relocatable_data")
artifact_toml_path = joinpath(@__DIR__() |> dirname, "Artifacts.toml")
tar_path = joinpath(@__DIR__(), "xlsx_artifacts.tar.gz")

@assert isdir(artifact_content)
package(artifact_content, tar_path)

add_artifact!(
    artifact_toml_path,
    "XLSX_relocatable_data",
    "https://www.dropbox.com/s/9rgm3pnd1dmzmy9/xlsx_artifacts.tar.gz?dl=1";
    force = true,
    lazy = false
)