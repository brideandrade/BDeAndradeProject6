import openpyxl
import numbers

def main():
    game_worksheet = open_worksheet("games-features.xlsx")
#Go through and change these if time for easier answers
    run_fxn_answers = ["1", "2", "3", "4", "5", "6"]
    while True:
        response = input("What command would you like to see? Type the corresponding number next to the command you'd like.\n"
          "0. Program quit? 1. Find the most players? 2.Find age restricted games? \n"
          "3. Find often recommended games? 4.Find well played games? \n"
          "5. See a count for the number of games? Or, 6. Look up more data?\n")
        response = response.lower()
        if response not in run_fxn_answers: break

        if response == "1":
            find_Most_Players(game_worksheet)

        if response == "2":
            find_Age_Restricted(game_worksheet)

        if response == "3":
            find_Often_Recommended(game_worksheet)

        if response == "4":
            find_well_played(game_worksheet)

        if response == "5":
            find_Count_For_System(game_worksheet)

        if response == "6":
            Look_Up_Data(game_worksheet)

def open_worksheet(file):
    game_excel = openpyxl.load_workbook(file)
    game_sheet = game_excel.active
    return game_sheet

def find_Most_Players(game_worksheet):
    largest_row = None
    print("You chose to Find Most Players")
    fmp_inp = input("Which system, Windows, Mac, or Linux, do you want to consider?")
    if fmp_inp == "windows":
        for row in game_worksheet.rows:
            windows_platform = row[26]
            windows_platform_value = windows_platform.value
            if windows_platform_value == "True":
                if largest_row is None:
                    largest_row = row
                if largest_row[17].value < row[17].value:
                    largest_row = row
        print(f"Windows: Game title: {largest_row[2].value}, Release date: {largest_row[4].value}, Player Count: {largest_row[17].value}")
        return
    if fmp_inp == "mac":
        for row in game_worksheet.rows:
            mac_platform = row[28]
            mac_platform_value = mac_platform.value
            if mac_platform_value == "True":
                if largest_row is None:
                    largest_row = row
                if largest_row[17].value < row[17].value:
                    largest_row = row
        print(f"Mac: Game title: {largest_row[2].value}, Release date: {largest_row[4].value}, Player Count: {largest_row[17].value}")
        return
    if fmp_inp == "linux":
        for row in game_worksheet.rows:
                linux_platform = row[27]
                linux_platform_value = linux_platform.value
                if linux_platform_value == "True":
                    if largest_row is None:
                        largest_row = row
                    if largest_row[17].value < row[17].value:
                        largest_row = row
        print(f"Windows: Game title: {largest_row[2].value}, Release date: {largest_row[4].value}, Player Count: {largest_row[17].value}")
        return

def find_Age_Restricted(game_worksheet):
    print("You chose to Find Age Restricted Games")
    for row in game_worksheet.rows:
        age_restrictions = row[5].value
        response_name = row[3].value
        release_date = row[4].value

        if not isinstance(age_restrictions, numbers.Number):
            continue
        if age_restrictions >= 17:
            print(f"Name:{response_name}, Date of release: {release_date}")

def find_Often_Recommended(game_worksheet):
    print("You chose to Find Recommended Games")
    for_inp = int(input("What is the minimum number of recommendations you're looking for?"))
    for row in game_worksheet.rows:
        recommendation_count = row[12].value
        number_of_owners = row[15].value
        if not isinstance(recommendation_count, numbers.Number):
            continue
        if number_of_owners == 0:
            continue
        recommendation_percent = recommendation_count / number_of_owners
        game_name = row[2].value
#How do I get it so that the recommendation_count does not print 0 recommendations and only greater than or equal to user input
#Same question for next fxn
        if for_inp > 0:
            print(f"Name:{game_name}, Recommendation Count:{recommendation_count}, Number of owners:{number_of_owners}, Percent of Recommendation:{recommendation_percent}")

def find_well_played(game_worksheet):
    print("You chose to Find Well Played Games")
    fwp_answer = int(input("What should the cutoff percentage be?"))

    for row in game_worksheet.rows:
        num_of_players = row[17].value
        num_of_owners = row[15].value
        if not isinstance(num_of_players, numbers.Number):
            continue
        if num_of_owners == 0:
            continue
        fwp_percent = num_of_players / num_of_owners
        game_title = row[2].value
        if fwp_percent > fwp_answer:
            print(f"Percentage: {fwp_percent}, Game Title: {game_title}")

def find_Count_For_System(game_worksheet):
    print("You chose to see a count for the number of games")
    count = 0
    for row in game_worksheet.rows:
            mac_platform = row[28].value
            windows_platform = row[26].value
            linux_platform = row[27].value
            if windows_platform == "True":
                if not isinstance(windows_platform, numbers.Number):
                    continue
                count +=1
            print(f"The count for number of Windows games available: {count}")
            if mac_platform == "True":
                if not isinstance(mac_platform, numbers.Number):
                    continue
                count += 1
            print(f"The count for number of Mac games available: {count}")
            if linux_platform == "True":
                if not isinstance(linux_platform, numbers.Number):
                    continue
                count += 1
            print(f"The count for number of Linux games available: {count}")
    return count
#There is something wrong in the code that won't have this part run correctly

def Look_Up_Data(game_worksheet):
    print("You chose to look up more data")
    data_question = input("Which game would you like more data on?")
    data_question_value = data_question.value
    for row in game_worksheet.rows:
        mac_platform = row[28].value
        windows_platform = row[26].value
        linux_platform = row[27].value
        response_name = row[3].value
#Is there a way to print Platform name without imitating "if mac_platform == "True"..."?
#How do I get it so that no matter what the user types, the following code prints what directions want
        if data_question_value == response_name:
            print(f"Name: {response_name}")


main()

