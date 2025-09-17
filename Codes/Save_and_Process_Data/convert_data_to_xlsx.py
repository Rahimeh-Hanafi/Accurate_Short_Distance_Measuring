import pandas as pd

# Read the text file
with open("data.txt", "r") as f:
    lines = f.readlines()

# Prepare data
blocks = []
current_block = []

for line in lines:
    line = line.strip()
    if not line or "---------------------------" in line:
        if current_block:
            blocks.append(current_block)
            current_block = []
    else:
        # Remove "MeasuredRange:" prefix and convert to float
        value = float(line.replace("MeasuredRange:", ""))
        current_block.append(value-1.15) # 1.15 is the offset of reading

# Add the last block if it wasn't added
if current_block:
    blocks.append(current_block)

# Convert to DataFrame (each block is a column)
df = pd.DataFrame({f"{(i+1)*10}mm": block for i, block in enumerate(blocks)})

# Save to Excel
df.to_excel("MeasuredRanges.xlsx", index=False)

print("✅ Data saved to MeasuredRanges.xlsx")
