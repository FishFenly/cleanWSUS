# cleanWSUS

PowerShell scripts to connect to the UpdateServices API and perform cleanup actions.

-Wsus Decline-
Finds all updates,
Selects updates that are superseded and not declined,
Takes one parameter ExclusionPeriod to allow you to specify how old the updates you want to decline should be.

-Wsus Clean-
Runs tasks that the inbuilt cleanup wizard would run
