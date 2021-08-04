# ToTheReader
Code used to gather data for the 'To The Reader' database.

The core script, 'data gatherer', trawls through the EEBO-TCP P4 XML corpus to identify texts with the division type attributes listed in the script. It fills out an Excel spreadsheet ('TTR') with certain descriptive metadata for these texts, and outputs the division element as a separate XML file.

Additional scripts in Python notebooks import pre-existing data to the spreadsheet from external sources:

  'TTR Professional Non-professional play.ipynb' matches the playbooks in TTR with those in the Database of Early English Playbooks (DEEP) by their STC number. It marks in TTR whether the playbook is a professional or non-professional play.
  'TTR Genre.ipynb' imports the genre categories assigned by Alan Farmer and Zachary Lesser into TTR.

The following scripts clean and standardize information from EEBO and the ESTC and join it to TTR:

  'TTR Format.ipynb' converts the 'physical description' field of the ESTC data into a standard format ('quarto', 'folio', etc).
  'TTR Author.ipynb' takes information held in the <signed> element of the EEBO-TCP files and cleans/standardizes it into set of author names.
  
The following scripts create new data and join it to TTR:
  
  'TTR Language.ipynb' identifies the language of each individual text.
  'TTR Category.ipynb' sorts the texts into one of twelve categories by their division type.
