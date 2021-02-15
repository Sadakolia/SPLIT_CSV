# SPLIT_CSV
Help to Split any CSV Data File into number of multiple small CSV Files


        .SYNOPSIS
        This is a Simple Script to Split any CSV file with three different Options
        	1) Split CSV with Number of Record Set.
            2) Split CSV with Number of Batches.
            3) Split CSV with Alphabets.
                    
        .DESCRIPTION
        Once you run the Script it will open "File Open Dialog Box " once you select the file it will give you below Options.
        	1) Split with Number of Record Set.
            2) Split in Batches.
            3) Split with Alphabets.
        
        1) Split with Number of Record Set.
            # As the Name indicates this option will be use to split the Selected File with number of Record Set available.
            
        2) Split in Batches
            # This option will split the CSV into Number of Batches as mention by the User

        3) Split with Alphabets (It is NOT Case Sensitive.
            # This Option will Split the CSV with Single or Range of Alphabets. 
            For Example : 
                        * A
                        * A-C
                        * E-Z
                        * A-Z

        .INPUTS
        None.

        .OUTPUTS
        None.

        .EXAMPLE
        1) Split with Number of Record Set.
        If Total Records in given CSV is 4647 then..

        Please Enter the Number of RecordSet to Split the CSV : 600
        Number of CSV Files created will be : 8
        Current Folder Path : "It will show you the current Primary CSV Path"

        .EXAMPLE
        2) Split with Number of Record Set.
        If Total Records in given CSV is 4647 then..

        Please Enter the Number of Batches to Split the CSV : 4
        Number of CSV Files created will be : 4
        Current Folder Path : "It will show you the current Primary CSV Path"

        .EXAMPLE
        3) Split with Alphabets.
        The File can be split with Single or Range of Alphabets. 
        Sample :-
                * A
                * A-C
                * E-Z
                * A-Z

        .LINK
        None

        .NOTES
        Author : Santosh Dakolia (Santosh.Dakolia@Microsoft.com)
        Last Edit : 15 Feb 2021
        Version 1.0 - Initial Release
