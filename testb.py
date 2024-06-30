import re

def split_text_by_uppercase(text):
    """
    Splits the text by uppercase words which are assumed to be speaker names,
    avoiding splitting character names like "DR WALSH".

    :param text: The text to split.
    :return: A list of strings, each starting with a speaker name.
    """
    # Split the text by spaces to get words
    words = text.split()
    
    # Find indices of all uppercase words longer than one character
    uppercase_indices = [i for i, word in enumerate(words) if word.isupper() and len(word) > 1]
    
    # Remove indices if the previous word is also uppercase
    filtered_indices = []
    for i in range(len(uppercase_indices)):
        if i == 0 or (uppercase_indices[i] - 1 != uppercase_indices[i - 1]):
            filtered_indices.append(uppercase_indices[i])
    
    print(words)
    print(uppercase_indices)
    print(filtered_indices)

    # Create segments based on the filtered indices
    segments = []
    start_index = 0
    for index in filtered_indices:
        segment = " ".join(words[start_index:index])
        if segment:  # Add non-empty segments
            segments.append(segment.strip())
        start_index = index
    # Add the last segment
    segments.append(" ".join(words[start_index:]).strip())
    
    return segments

# Example usage
text = ("LARA What garbage? STEVEN I don't know, minerals and shit. Micro chips. "
        "LARA Okay, yeah we'll, we'll pick one up tomorrow and we'll get some tin foil to block out the secret government signals too, yeah? "
        "STEVEN Right. " 
        "RADIO ANNOUNCER I am you. "
        "LARA I'm turning this shit off.")

split_text = split_text_by_uppercase(text)
for part in split_text:
    print(f"'{part}'")