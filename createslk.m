let
    CreateSLK = (Surname as text, Firstname as text, DOB as text, Sex as text, optional #"keep / in date" as logical) as text =>
    let
        //get a list of lists containing the numbers of Unicode punctuation characters
        numberlists = {{0..64},{91..96},{123..500}},
        //turn this into a single list
        combinedlist = List.Combine(numberlists),
        //get a list of all the punctuation characters that these numbers represent
        punctuationlist = List.Transform(combinedlist, each Character.FromNumber(_)),
        //remove these characters from Firstname and Surname
        alphaFirstname = Text.Remove(Firstname, punctuationlist),
        alphaSurname = Text.Remove(Surname, punctuationlist),
        //get letters 2,3 from Firstname using SLK581 standard
        lenFirstname = Text.Length(alphaFirstname),
        SLKFirstname = if lenFirstname > 2 then
                            Text.Middle(alphaFirstname,1,2)
                       else if lenFirstname > 1 then
                            Text.Middle(alphaFirstname,1,1) & "2"
                       else if lenFirstname = 1 then
                            "22"
                       else "99",
        upperFirstname = Text.Upper(SLKFirstname),
        //get letters 2,3 and 5 from Surname using SLK581 standard
        lenSurname = Text.Length(alphaSurname),
        SLKSurname = if lenSurname > 4 then
                        Text.Middle(alphaSurname,1,2) & Text.Middle(alphaSurname,4,1)
                     else if lenSurname > 2 then
                        Text.Middle(alphaSurname,1,2) & "2"
                     else if lenSurname > 1 then
                            Text.Middle(alphaSurname,1,1) & "22"
                     else if lenSurname = 1 then
                        "222"
                     else "999",
        upperSurname = Text.Upper(SLKSurname),
        //format DOB with leading zero and deal with /
        repairDOB = if Text.Length(DOB) = 9 then 
                        "0" & DOB
                    else if Text.Length(DOB) = 10 then
                        DOB
                    else "01/01/1900",
        SLKDOB = if #"keep / in date" = false then Text.Remove(repairDOB,{"\","/"}) else repairDOB ,
        //change Sex value to 1,2 or 9 if not already
        Sexformatted = if Sex = "1" then "1" else if Sex = "2" then "2" else if Text.StartsWith(Sex, "M") then "1" else if Text.StartsWith(Sex, "F") then "2" else "9",
        //create SLK581
        SLK581 = upperSurname & upperFirstname & SLKDOB & Sexformatted

    in
        SLK581
in
    CreateSLK