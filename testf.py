import re

def count_pattern_occurrences(file_path, encoding):
    # Define the regex pattern
    pattern = re.compile(r'\d{2}:\d{2}:\d{2}:\d{2}\s+--:--:--:--\s+.+')
    pattern = re.compile(r'\d{2}:\d{2}:\d{2}:\d{2}\s+--:--:--:--\s*')
    pattern = re.compile(r'(\d{2}:\d{2}:\d{2}:\d{2}\s+--:--:--:--\s*)|(\d{2}:\d{2}:\d{2}:\d{2}\s+\d{2}:\d{2}:\d{2}:\d{2}\s*)')
   
    # Read the content of the file
    with open(file_path, 'r', encoding=encoding) as file:
        content = file.read()
    
    # Find all matches in the text
    matches = pattern.findall(content)
      
    # Print each match
    for match in matches:
        print(match)

    # Return the count of matches
    return len(matches)
def extract_dialogs(file_path, encoding):
    # Define the regex pattern
    pattern = re.compile(r'\d{2}:\d{2}:\d{2}:\d{2}\s+(?:--:--:--:--|\d{2}:\d{2}:\d{2}:\d{2})\s+(.+)')
    
    # Read the content of the file
    with open(file_path, 'r', encoding=encoding) as file:
        content = file.read()
    
    # Find all matches in the text
    matches = pattern.findall(content)
    
    # Return the list of extracted dialogs
    return matches
# Example usage
file_path = 'example.txt'  # Path to your text file
encoding = 'utf-8'  # Encoding of the text file
# Example usage
file_path = 'example.txt'  # Path to your text file
encoding = 'utf-8'  # Encoding of the text file

# You can create an example text file with the provided content:
example_text = """02:59:54:13\t--:--:--:--\tWhere's Jake?
02:59:57:03\t02:59:59:14\tTying up some loose ends.
03:00:04:00\t03:00:06:10\tI'm gonna miss you.
03:00:09:14\t03:00:10:20\tThanks.
03:00:17:11\t03:00:19:04\tSaying good-bye?
03:00:20:10\t--:--:--:--\tGonna be a long ride.
03:00:22:01\t03:00:24:03\t- Travis: Sure was.\t- Three whiskeys.
03:00:26:00\t--:--:--:--\tI thought you two wanted\tto settle down someplace.
03:00:28:12\t--:--:--:--\tAll this got me thinking\tabout my own brother.
03:00:30:18\t--:--:--:--\tThought I better check in.
03:00:32:04\t--:--:--:--\tBesides,\twe better get out of here
03:00:34:20\t03:00:36:14\tbefore you get yourself\ta real preacher.
03:00:38:00\t--:--:--:--\tTurn this into\ta respectable town, right?
03:00:40:04\t--:--:--:--\tWell, that's the plan.
03:00:42:00\t--:--:--:--\tThere may be others out there.
03:00:43:20\t--:--:--:--\tI'll take care of them.\tLegally.
03:00:46:01\t--:--:--:--\tI still have\tfriends back east.
03:00:48:04\t03:00:50:15\tYou have friends\tout here, too.
03:00:52:05\t03:00:54:04\tThank you.
03:00:57:06\t--:--:--:--\tYou're always welcome here.
03:00:58:22\t03:01:01:00\tFriends.
03:01:07:14\t--:--:--:--\tYou really think\tthis town is gonna survive?
03:01:12:05\t--:--:--:--\tI got in touch\twith the railroad.
03:01:13:15\t03:01:15:12\tThey're gonna come here\tand take a look around.
03:01:16:19\t03:01:18:13\tI think we'll be all right.
03:01:27:15\t03:01:30:00\tYeah, I think\tyou'll be just fine.
03:01:32:20\t03:01:35:02\tThank you for everything.
03:01:36:15\t--:--:--:--\tMiss Josie.
03:01:38:04\t03:01:39:15\tCame to say good-bye.
03:01:42:21\t--:--:--:--\tYou take care of Jim now.
03:01:45:04\t03:01:47:15\t- I will.\t- He's a wild one.
03:01:53:04\t--:--:--:--\t( clicks tongue )
03:01:54:20\t--:--:--:--\tYou two try\tto stay out of trouble.
03:01:56:17\t--:--:--:--\t- We always try to.\t- Jim?
03:02:00:05\t--:--:--:--\tYou gonna run the saloon now?
"""

with open(file_path, 'w', encoding=encoding) as file:
    file.write(example_text)

# Test the function
#dialogs = count_pattern_occurrences(file_path, encoding)
#print('Extracted dialogs:'+str(dialogs)+" /"+str(len(example_text.split("\n"))))

dialogs = extract_dialogs(file_path, encoding)
print('Extracted dialogs:')
for dialog in dialogs:
    print(dialog)