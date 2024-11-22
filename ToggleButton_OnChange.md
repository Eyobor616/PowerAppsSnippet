
If(varLoadGroups = true,

        Set(varRuns, 0);
        //***** LOAD FIRST BATCH
        Set(
            varGroups,
            Office365Groups.HttpRequest(
                "https://graph.microsoft.com/v1.0/groups",
                "GET",
                ""
            )
        );
        ClearCollect(
            colGroups,
            Table(varGroups.value)
        );

        //***** START DO UNTIL LOOP

        ForAll(Sequence(10),
            If(! IsBlank(varGroups.'@odata.nextLink'),
            Select(btnCodeRepeat) //Parent.OnLoad_Groups()
            ));


        //****** END LOOP END

        If(
            IsBlank(varGroups.'@odata.nextLink'),

                Notify("Done", NotificationType.Success, 500);


                ClearCollect(
                colDistributionList_Group,
                ForAll(Table(
                Filter(        colGroups,        IsBlank(First(Value.groupTypes)),        Value.mailEnabled,        !Value.securityEnabled    )),
                {
                    displayName: Text(ThisRecord.Value.displayName),
                    id: Text(ThisRecord.Value.id),
                    description: Text(ThisRecord.Value.description),
                    //groupTypes: Text(ThisRecord.Value.groupTypes),
                    mail: Text(ThisRecord.Value.mail),
                    mailEnabled: Text(ThisRecord.Value.mailEnabled),
                    onPremisesDomainName: Text(ThisRecord.Value.onPremisesDomainName),
                    securityEnabled: Text(ThisRecord.Value.securityEnabled)
                }
            ));

            Set(varLoadGroups, false)


        )
);

