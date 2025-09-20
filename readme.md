arduino-cli compile --fqbn arduino:avr:nano:cpu=atmega328old -e BlinkNano
arduino-cli upload -p COM7 --fqbn arduino:avr:nano:cpu=atmega328old BlinkNano
arduino-cli board list
arduino-cli core install arduino:avr
