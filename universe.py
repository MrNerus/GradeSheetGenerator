COL_GREEN = '\033[38;5;46m'  # Green Color
COL_CYAN = '\033[38;5;51m'  # Cyan Color
COL_YELLOW = '\033[38;5;226m'  # Yellow Color
COL_RED = '\033[38;5;196m'  # Red Color
COL_RESET = '\033[0m'  # Color Reset      

# Capitalizes First Letter
def capitalizeFirst(string):
    stringL = string.split(" ")
    string2 = ""
    for i in range(0, len(stringL)):
        buffer = ""
        for j in range(0,len(stringL[i])):
            buffer += stringL[i][j].upper() if j == 0 else stringL[i][j]
        stringL[i] = buffer
        string2 += f"{stringL[i]} " if i != (len(stringL)-1) else f"{stringL[i]}"
    return string2

# If Api Receives "None", it converts to ""
def noneFilter(string):
    if string != "None":
        return string
    return ""
