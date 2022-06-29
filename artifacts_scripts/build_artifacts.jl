using ArtifactUtils

artifact_content = joinpath(@__DIR__() |> dirname, "relocatable_data")
@assert isdir(artifact_content)

artifact_id = artifact_from_directory(artifact_content)
# For this to work we have to be logged in Github with the Github CLI
gist = upload_to_gist(artifact_id)
add_artifact!(joinpath(@__DIR__() |> dirname, "Artifacts.toml"), "relocatable_data", gist)
