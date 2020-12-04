import json

class Ripper:

    def __init__(self, path):
        self.path = path
        self.entries = {}

    def load(self):
        """
            Loads and builds the ripper data set
        """
        print(f"Loading ripper file: {self.path}")
        
        # Load ripper file into memory
        f = open(self.path)

        # Parse JSON data from file
        self.__build(json.load(f))
        
        for key, value in self.entries.items() :
            print(key)

        e = self.entries["SS191/SF191/NS191"]
        print(json.dumps(e))

        # There are 318 subject offerings
        f.close()
    
    def __build(self, json_data):
        print(f"Building data from {len(json_data)} entries")

        for entry_data in json_data:
            section = entry_data["section"]
            
            if section not in self.entries:
                self.entries[section] = []
            
            # Build the subject entry object
            sub_entry = {}
            subject_data = entry_data["subject"]
            sub_entry["code"] = subject_data["code"]
            sub_entry["name"] = subject_data["name"]

            # Create detail data            
            detail_data = entry_data["subject_offering_details"]
            det_entry = {}
            for detail in detail_data:
                det = {}
                day = detail["day_of_weeks"]["day_string"]

                # Some subjects do not have rooms
                room = "N/A" if detail["rooms"] is None else detail["rooms"]["code"]
                
                det["time_start"] = detail["time_start"]
                det["time_end"] = detail["time_end"]
                det["room"] = room

                # Append detail data to entry
                det_entry[day] = det
            
            sub_entry["schedules"] = det_entry

            # Append created subject entry object
            self.entries[section].append(sub_entry)


if __name__ == "__main__":
    r = Ripper("subjectoffering.json")
    r.load()