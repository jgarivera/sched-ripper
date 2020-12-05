from ripper import Ripper
from excella import Excella

if __name__ == "__main__":
    r = Ripper("subjectoffering.json")
    r.load()
    
    entries = r.get_entries()

    # Render ripper entries into Excel spreadsheet
    e = Excella(entries, "scheds.xlsx") 
    e.begin()
    e.close()