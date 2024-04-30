text = """Come on... hurry up.
No!
Sit! What? What?
Come on, hurry up, you piece of shit.
Just take him with you, let him run around.
Shut up.
Shut up. Go.
What?
Shut up.
Shut up. Shut up!
Come on.
Shut up!
What are you looking at?"""

# Split the text into lines
lines = text.split('\n')

# Iterate through each line and count the characters
character_counts = {line: len(line) for line in lines}

# Print the results
charactercount=0
linecount=0
for line, count in character_counts.items():
    print(f"'{line}' ->  {count} caractères")
    charactercount=charactercount+count
    linecount=linecount+1

print("-----------------------------")
print("Caractères : "+str(charactercount))
print("Lignes     : "+str(linecount))
print("Repliques  : ",str(charactercount/40))