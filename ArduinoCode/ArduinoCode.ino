#define Volume1_Pin A0
#define Volume2_Pin A1
#define Volume3_Pin A2
#define Volume4_Pin A3
#define Volume5_Pin A4

#define Mute1_Pin 2
#define Mute2_Pin 3
#define Mute3_Pin 4
#define Mute4_Pin 5
#define Mute5_Pin 6

bool Muted1 = false;
bool Muted2 = false;
bool Muted3 = false;
bool Muted4 = false;
bool Muted5 = false;

unsigned long LastButtonDebounceTime = 0;
unsigned long ButtonDebounceTime = 300;
unsigned long LastSeialDebounceTime = 0;
unsigned long SerialDebounceTime = 50;

void setup()
{
  Serial.begin(115200);
}

void loop()
{
  if (millis() - LastButtonDebounceTime > ButtonDebounceTime)
  {
    if (digitalRead(Mute1_pin) == LOW)
    {
      if (Muted1 == false)
      {
        Muted1 = true;
      }
      else
      {
        Muted1 = false;
      }
    }
    if (digitalRead(Mute2_pin) == LOW)
    {
      if (Muted2 == false)
      {
        Muted2 = true;
      }
      else
      {
        Muted2 = false;
      }
    }
    if (digitalRead(Mute3_pin) == LOW)
    {
      if (Muted3 == false)
      {
        Muted3 = true;
      }
      else
      {
        Muted3 = false;
      }
    }
    if (digitalRead(Mute4_pin) == LOW)
    {
      if (Muted4 == false)
      {
        Muted4 = true;
      }
      else
      {
        Muted4 = false;
      }
    }
    if (digitalRead(Mute5_pin) == LOW)
    {
      if (Muted5 == false)
      {
        Muted5 = true;
      }
      else
      {
        Muted5 = false;
      }
    }
    LastButtonDebounceTime = millis();
  }

  int volume1 = (analogRead(Volume1_Pin) / 10);
  int volume2 = (analogRead(Volume2_Pin) / 10);
  int volume3 = (analogRead(Volume3_Pin) / 10);
  int volume4 = (analogRead(Volume4_Pin) / 10);
  int volume5 = (analogRead(Volume5_Pin) / 10);

  if (millis() - LastSeialDebounceTime > SerialDebounceTime)
  {
    Serial.print("Volume1:");
    Serial.print(volume1);
    Serial.print(", Volume2:");
    Serial.print(volume2);
    Serial.print(", Volume3:");
    Serial.print(volume3);
    Serial.print(", Volume4:");
    Serial.print(volume4);
    Serial.print(", Volume5:");
    Serial.print(volume5);
    Serial.print(", Mute1:");
    Serial.print(Muted1);
    Serial.print(", Mute2:");
    Serial.print(Muted2);
    Serial.print(", Mute3:");
    Serial.print(Muted3);
    Serial.print(", Mute4:");
    Serial.print(Muted4);
    Serial.print(", Mute5:");
    Serial.println(Muted5);
    
    LastSerialDebounceTime = millis();
  }
}
