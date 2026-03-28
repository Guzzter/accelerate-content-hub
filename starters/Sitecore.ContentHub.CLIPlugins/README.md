# Content Hub CLI Plugins Starter

This is a starter project intended to help in the development of CLI plugins for Content Hub.

> ⚠️ **Note**
>
> The provided code is intended as a guideline and must be tailored to suit your specific implementation requirements. Please ensure thorough end-to-end testing is conducted to validate its functionality and performance in your environment.

## Develop

The solution contains two projects `Sitecore.CH.Cli.Plugins.Base` and `Sitecore.CH.Cli.Plugins.Starter`.

The base project contains services which will be shared between plugins.

The starter project contains some commands to demonstrate how plugins can be developed.

## Debugging

In the `Sitecore.CH.Cli.Plugins.Starter` project, run `Init-For-Debug.ps1` - this must be run as an administrator. This creates a symlink in your CH CLI plugins folder to a new output directory for the the starter plugin.

Debug profiles have been created for all of the commands in the starter plugin, allowing commands to be debugged within Visual Studio.
