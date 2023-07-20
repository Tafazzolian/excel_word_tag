import pandas as pd
import re

# read in the Excel file
df = pd.read_excel('input.xlsx')

# iterate over each row in the DataFrame
for index, row in df.iterrows():
    # extract the word from the second row and make it lowercase
    word = str(row[1]).lower()
    
    # extract the sentence from the sixth row
    sentence = str(row[5])
    
    # use regular expressions to find the word in the sentence
    pattern = r"\b{}\b".format(re.escape(word))
    match = re.search(pattern, sentence, re.IGNORECASE)
    
    if match:
        # tag the matched word with <>
        new_sentence = re.sub(pattern, f"<{match.group()}>", sentence, flags=re.IGNORECASE)
        
        # update the DataFrame with the new sentence
        df.at[index, 5] = new_sentence
    else:
        # print out the row index for each row that is skipped
        print(f"Row {index} skipped: word '{word}' not found in sentence '{sentence}'")
        
# write the updated DataFrame to a new Excel file
df.to_excel('output.xlsx', index=False)
