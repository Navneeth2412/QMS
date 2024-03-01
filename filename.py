import os
import datetime

def get_next_letter(current_letter):
    if current_letter == 'Z':
        return 'A'
    else:
        return chr(ord(current_letter) + 1)

def create_file():
    current_date = datetime.date.today().strftime("%y%m%d")
    
    if not os.path.exists('last_state.txt'):
        with open('last_state.txt', 'w') as f:
            f.write(f"{current_date}_A")
        return f"{current_date}_A"

    with open('last_state.txt', 'r') as f:
        last_state = f.read().strip()

    last_date, last_letter = last_state.split('_')
    
    if last_date == current_date:
        next_letter = get_next_letter(last_letter)
    else:
        next_letter = 'A'

    new_state = f"{current_date}_{next_letter}"

    with open('last_state.txt', 'w') as f:
        f.write(new_state)

    return new_state

if __name__ == "__main__":
    new_file_name = create_file()
    print(f"Created new file: {new_file_name}")
