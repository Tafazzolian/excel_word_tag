import pandas as pd
import re

df = pd.read_excel('input.xlsx')

for index, row in df.iterrows():
    word = str(row[1]).lower()
    
    sentence = str(row[5])
    
    pattern = r"\b{}\b".format(re.escape(word))
    match = re.search(pattern, sentence, re.IGNORECASE)
    
    if match:
        new_sentence = re.sub(pattern, f"<{match.group()}>", sentence, flags=re.IGNORECASE)
        
        df.at[index, 5] = new_sentence
    else:
        print(f"Row {index} skipped: word '{word}' not found in sentence '{sentence}'")
        
df.to_excel('output.xlsx', index=False)
