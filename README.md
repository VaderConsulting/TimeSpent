# TimeSpent

TimeSpent is a C# console app that exports Exchange calendar appointments to CSV so you can review hours spent from Outlook. You pass `/start`, `/end`, `/mailpath` (EWS URL), optional `/domain` `/username` `/password`, `/version` (2007 through 2010_SP2), and `/out`; it FindItem-s the Calendar folder and writes Subject,Date,Start,End. `TimeSpentLib` wraps Exchange Web Services (`ExchangeServiceBinding` / EWS.dll) and scopes dates to W. Australia Standard Time. This tree is Dave Robinson's working copy of the Microsoft IT MSDN sample (archive.msdn.microsoft.com/timespent, MS-PL, copyright Microsoft IT 2009).

**Source last updated:** 2013-12-09 · **Language:** C# · **Target:** .NET Framework 4.0 (Client Profile on the console) · **Output:** console exe (`TimeSpent.exe`) + class library

## Solution structure

| Project | Language | Type | Purpose |
|---------|----------|------|---------|
| `TimeSpent` | C# | Console exe (.NET 4.0 Client) | Parses switches and writes calendar rows to CSV or stdout |
| `TimeSpentLib` | C# | Class library (.NET 4.0) | EWS connect + `GetCalendarItems` (`CalendarItemData` / `CalendarItemList`) |
| `TimeSpentModel` | UML / layer diagram | VS Team Architect modeling project | Layer diagram (console vs library); not in `TimeSpent.sln` |

## How to open

Open `TimeSpent.sln` in Visual Studio Express 2013 for Windows Desktop (solution format 12.00). The C# projects are ToolsVersion 4.0 (`ProductVersion` 8.0.30703). `TimeSpentLib` expects `EWS.dll` at `..\..\EWS.dll` relative to the library project and `Microsoft.Exchange.WebServices.dll` from the EWS Managed API 2.0 install (`Program Files\Microsoft\Exchange\Web Services\2.0\`). Debug start arguments live in `TimeSpent/TimeSpent.csproj.user` (gitignored); copy `TimeSpent/TimeSpent.csproj.user.example` and fill in your EWS URL and credentials. The modeling project `TimeSpentModel/TimeSpentModel.modelproj` is on disk but is not included in the solution.

## Requirements

- Visual Studio 2013, .NET Framework 4.0

## Attribution and provenance

Working copy from Dave Robinson's OneDrive Historical Dev folder `TimeSpent`. Assembly title/product `TimeSpent` / `TimeSpentLib`, company Microsoft IT, copyright © Microsoft IT 2009. Source comments point at `http://archive.msdn.microsoft.com/timespent` and the MS-PL. Dave's copy adds credentialed EWS connect, Exchange 2010 family versions, and AWST timezone scoping. A usage-example EWS hostname in `CmdLine.cs` was replaced with `https://exchange.example/EWS/Exchange.asmx`. Debug start arguments that held a domain, username, and password are not in git (`*.user`); see the `.example` file.

## License

Original Microsoft Public License (MS-PL) terms apply to the Microsoft IT sample. See `LICENSE` and `THIRD_PARTY_NOTICES.md`. Dave Robinson's working-copy changes in this catalogue are also offered under MS-PL so the tree stays under one license. Exchange Web Services proxy types (`ExchangeWebServices`) and the EWS Managed API are Microsoft Exchange artifacts.
