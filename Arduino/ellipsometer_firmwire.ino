#include <util/delay.h>
#include <stdlib.h>

char str[32];
const int delay1 = 25;
const int delay2 = 10;
const int pwmPin = 5;
int pwmValue;

bool dirHigh;
void setup() {
  DDRB |= 0b11110000;
  DDRH |= 0b01111000;
  setupDRV8825();
  pinMode(pwmPin, OUTPUT);
  Serial.begin(9600);
}

void setupDRV8825()
{
  DDRL |= 0b01010101;
  DDRG |= 0b00000101;
  DDRC |= 0b00000101;
  PORTL |= (1 << 4) | (1 << 6);
  PORTG |= (1 << 2);
}

void moveStep( int steps, byte dir) {
  if (dir == 1) {
    PORTL |= (1 << 0);
  } else {
    PORTL &= ~(1 << 0);
  }
  for (int i = 0; i < steps; i++) {
    PORTL |= (1 << 2);
    delay(15);
    PORTL &= ~(1 << 2);
    delay(15);
  }
}

//барабан
void moveBackward1(int steps) {
  for (int i = 0; i < steps; i++) {
    PORTH = (1 << 6);
    _delay_ms(delay1);
    PORTH = (1 << 5);
    PORTH = (1 << 6);
    _delay_ms(delay1);
    PORTH = (1 << 5);
    _delay_ms(delay1);
    PORTH = (1 << 5);
    PORTH = (1 << 4);
    _delay_ms(delay1);
    PORTH = (1 << 4);
    _delay_ms(delay1);
    PORTH = (1 << 4);
    PORTH = (1 << 3);
    _delay_ms(delay1);
    PORTH = (1 << 3);
    _delay_ms(delay1);
    PORTH = (1 << 3);
    PORTH = (1 << 6);
    _delay_ms(delay1);
  }
}

// анализатор
void moveBackward2(int steps) {
  for (int i = 0; i < steps; i++) {
    PORTB = (1 << 7);
    _delay_ms(delay2);
    PORTB = (1 << 6);
    PORTB = (1 << 7);
    _delay_ms(delay2);
    PORTB = (1 << 6);
    _delay_ms(delay2);
    PORTB = (1 << 6);
    PORTB = (1 << 5);
    _delay_ms(delay2);
    PORTB = (1 << 5);
    _delay_ms(delay2);
    PORTB = (1 << 5);
    PORTB = (1 << 4);
    _delay_ms(delay2);
    PORTB = (1 << 4);
    _delay_ms(delay2);
    PORTB = (1 << 4);
    PORTB = (1 << 7);
    _delay_ms(delay2);
  }
}

void moveForward1(int steps) {
  for (int i = 0; i < steps; i++) {
    PORTH = (1 << 3);
    PORTH = (1 << 6);
    _delay_ms(delay1);
    PORTH = (1 << 3);
    _delay_ms(delay1);
    PORTH = (1 << 4);
    PORTH = (1 << 3);
    _delay_ms(delay1);
    PORTH = (1 << 4);
    _delay_ms(delay1);
    PORTH = (1 << 5);
    PORTH = (1 << 4);
    _delay_ms(delay1);
    PORTH = (1 << 5);
    _delay_ms(delay1);
    PORTH = (1 << 5);
    PORTH = (1 << 6);
    _delay_ms(delay1);
    PORTH = (1 << 6);
    _delay_ms(delay1);
  }
}

void moveForward2(int steps) {
  for (int i = 0; i < steps; i++) {
    PORTB = (1 << 4);
    PORTB = (1 << 7);
    _delay_ms(delay2);
    PORTB = (1 << 4);
    _delay_ms(delay2);
    PORTB = (1 << 5);
    PORTB = (1 << 4);
    _delay_ms(delay2);
    PORTB = (1 << 5);
    _delay_ms(delay2);
    PORTB = (1 << 6);
    PORTB = (1 << 5);
    _delay_ms(delay2);
    PORTB = (1 << 6);
    _delay_ms(delay2);
    PORTB = (1 << 6);
    PORTB = (1 << 7);
    _delay_ms(delay2);
    PORTB = (1 << 7);
    _delay_ms(delay2);
  }
}

void readSerial() {
  byte buffer = 0;
  memset(str, 0, 32);
  if (Serial.available()  > 0) {
    delay(250);
    buffer = Serial.available();
    for (int i = 0; i < buffer; i++)
    {
      char c = Serial.read();
      str[i] = c;
    }
  }
}

void setVoltage() {
  pwmValue = 0;
  analogWrite(pwmPin, pwmValue);
}

void increaseVoltage() {
  if (pwmValue > 220) {
    pwmValue = 220;
  }
  analogWrite(pwmPin, pwmValue++);
}

void decreaseVoltage() {
  analogWrite(pwmPin, pwmValue--);
}


void loop() {
  readSerial();
  if (strcmp(str, "forward1") == 0) {
    moveStep(38, 1);
    Serial.println("GoForward1:");
  }

  if (strcmp(str, "forward2") == 0) {
    moveForward2(80);
    Serial.println("GoForward2:");
  }

  if (strcmp(str, "backward1") == 0) {
    moveStep(390, 0);
    Serial.println("GoBackward1:");
  }

  if (strcmp(str, "backward2") == 0) {
    moveBackward2(40);
    Serial.println("GoBackward2:");
  }

  if (strcmp(str, "setVoltage") == 0) {
    setVoltage();
    Serial.println("voltage is set up" + pwmValue);
  }

  if (strcmp(str, "increase") == 0) {
    increaseVoltage();
  }

  if (strcmp(str, "decrease") == 0) {
    decreaseVoltage();
  }

  if ( atol(str) > 0 && atol(str) < 255 ) {
    pwmValue = atol(str);
    analogWrite(pwmPin, pwmValue);
  }
}


