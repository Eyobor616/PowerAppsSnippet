If(
    !IsBlank(varGroups.'@odata.nextLink'),
    Set(varRuns,varRuns+1);
    //***** LOAD NEXT BATCH
    Set(varGroups,Office365Groups.HttpRequest(varGroups.'@odata.nextLink',"GET",""));
    Collect(colGroups,Table(varGroups.value))
);
