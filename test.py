class Character():
    MAX_SPEED = 100

    def __init__(self, race, damage=10):
        self.damage = damage
        self.race = race

        self._current_speed = 50

    @property
    def current_speed(self):
        return self._current_speed
    @current_speed.setter
    def current_speed(self, current_speed):
        if current_speed < 0:
            self._current_speed = 0
        elif current_speed > 100:
            self._current_speed = 100
        else:
            self._current_speed = current_speed

unit = Character('Ork')
unit._current_speed = 50
print(unit._current_speed)

