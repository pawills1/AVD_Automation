$modules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.DeviceManagement",
    "Microsoft.Graph.DeviceManagement.Administration",
    "Microsoft.Graph.Identity.DirectoryManagement",
    "Microsoft.Graph.Security",
    "Microsoft.Graph.Beta.DeviceManagement"
)

foreach ($module in $modules) {
    Install-Module -Name $module -RequiredVersion 2.25.0 -Scope AllUsers -Force -Verbose
}