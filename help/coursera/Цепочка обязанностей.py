class Character:
    def __init__(self):
        self.name = "Nagibator"
        self.xp = 0
        self.passed_quests = set()
        self.taken_quests = set()


def add_quest_speak(char):
    quest_name = "Поговорить с фермером"
    xp = 100
    if quest_name not in (char.passed_quests | char.taken_quests):
        print(f"Квест получен: \"{quest_name}\"")
        char.taken_quests.add(quest_name)
    elif quest_name in char.taken_quests:
        print(f"Квест сдан: \"{quest_name}\"")
        char.passed_quests.add(quest_name)
        char.taken_quests.remove(quest_name)
        char.xp += xp


def add_quest_hunt(char):
    quest_name = "Охота на крыс"
    xp = 300
    if quest_name not in (char.passed_quests | char.taken_quests):
        print(f"Квест получен: \"{quest_name}\"")
        char.taken_quests.add(quest_name)
    elif quest_name in char.taken_quests:
        print(f"Квест сдан: \"{quest_name}\"")
        char.passed_quests.add(quest_name)
        char.taken_quests.remove(quest_name)
        char.xp += xp


def add_quest_carry(char):
    quest_name = "Принести доски из сарая"
    xp = 200
    if quest_name not in (char.passed_quests | char.taken_quests):
        print(f"Квест получен: \"{quest_name}\"")
        char.taken_quests.add(quest_name)
    elif quest_name in char.taken_quests:
        print(f"Квест сдан: \"{quest_name}\"")
        char.passed_quests.add(quest_name)
        char.taken_quests.remove(quest_name)
        char.xp += xp


class QuestGiver:
    def __init__(self):
        self.quests = []

    def add_quest(self, quest):
        self.quests.append(quest)

    def handle_quests(self, character):
        for quest in self.quests:
            quest(character)


all_quests = [add_quest_speak, add_quest_hunt, add_quest_carry]

quest_giver = QuestGiver()

for quest in all_quests:
    quest_giver.add_quest(quest)

player = Character()

quest_giver.handle_quests(player)

print("Получено: ", player.taken_quests)
print("Сдано: ", player.passed_quests)
player.taken_quests = {'Принести доски из сарая', 'Поговорить с фермером'}
quest_giver.handle_quests(player)
print("Получено: ", player.taken_quests)
print("Сдано: ", player.passed_quests)

quest_giver.handle_quests(player)
print("Получено: ", player.taken_quests)
print("Сдано: ", player.passed_quests)
