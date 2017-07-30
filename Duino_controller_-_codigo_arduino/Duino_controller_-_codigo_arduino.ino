#include <DHT.h>

// variaveis

#define DHTTYPE DHT22
#define DHTPIN 7
int led1 = 13;
int led2 = 12;
int led3 = 11;
int led4 = 10;
int led5 = 9;
int analog = 0;
int analog1 = 0;

int analog2 = 0;
double voltage = 0;
double amps = 0;
int ACSoffset = 2500;
int mVperAmp = 66;

float temperatura;
float umidade;
int acionamentoPIR;

int pinopir = 6;
DHT dht(DHTPIN, DHTTYPE);

// configuração
void setup() {
Serial.begin(9600);
pinMode(led1, OUTPUT);
pinMode(led2, OUTPUT);
pinMode(led3, OUTPUT);
pinMode(led4, OUTPUT);
pinMode(led5, OUTPUT);

dht.begin();
}

void loop() {
char leitura = Serial.read();
//verifica se é um teste
if(leitura == 't')
{
  Serial.println("t");
  delay (50);
  Serial.println("Teste OK!!"); 
}
// envia leitura
else if (leitura == 'T')
{
  temperatura = dht.readTemperature();
  umidade = dht.readHumidity();
  Serial.println("T");
  Serial.println(temperatura);
  delay(50);
  Serial.println("H");
  delay (30);
  Serial.println(umidade);
}
//aciona relés
else if (leitura == '1')
{
  digitalWrite (led1, HIGH);
    
}
else if (leitura == '2')
{
  digitalWrite (led2, HIGH);
    
}
else if (leitura == '3')
{
  digitalWrite (led3, HIGH);
    
}
else if (leitura == '4')
{
  digitalWrite (led4, HIGH);
    
}
else if (leitura == '5')
{
  digitalWrite (led5, HIGH);
    
}
// PARA DESLIGAR RELÉS
else if (leitura == 'A')
{
  digitalWrite (led1, LOW);
    
}

else if (leitura == 'B')
{
  digitalWrite (led2, LOW);
    
}

else if (leitura == 'C')
{
  digitalWrite (led3, LOW);
    
}

else if (leitura == 'D')
{
  digitalWrite (led4, LOW);
    
}

else if (leitura == 'E')
{
  digitalWrite (led5, LOW);
    
}

// verifica o sensor de movimento
acionamentoPIR = digitalRead(pinopir); //Le o valor do sensor PIR
 if (acionamentoPIR == HIGH)  //Sem movimento, mantem led azul ligado
 {
    Serial.println("P");
    delay (15);
 }
// envia o sinal do LDR
analog = analogRead(0);
Serial.println("l");
delay (50);
Serial.println(analog);
delay (350);  

// envia o sinal do divisor de tensão para rede elétrica

analog1 = analogRead(1);
Serial.println("r");
delay (50);
Serial.println(analog1);
delay (200);

//calcula corrente

analog2 = analogRead(2);
  voltage = (analog2 / 1024.0) * 5000; // Gets you mV
  amps = ((voltage - ACSoffset) / mVperAmp);
  if (amps < 0)
  {
    amps = ((voltage - ACSoffset) / mVperAmp) * -1;
  }
  Serial.println("a");
  delay (50);
  Serial.println(amps);
  delay (50);
  
}
