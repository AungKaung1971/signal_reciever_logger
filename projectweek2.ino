#include <RadioHead.h>

// 
// Project Week 2: Wireless Communications 
//
// RFM69HCW receiver/transmitter code
// 
// Requires: Nano ESP32, I2C OLED display RFM69HCW radio
// Configured for Arduino Nano ESP32 (also tested on Due)
// 
// Developed with PlatformIO and Arduino framework
//
// Robert J Watson, January 2025
// Department of Electronic & Electrical Engineering 
// University of Bath
//

#include <Arduino.h>
#include <SPI.h>
#include <Adafruit_GFX.h>
#include <Adafruit_SSD1306.h>
#include <RadioHead.h>
#include <RH_RF69.h>

// This is likely the only #define that likely will change
#define RF69_FREQ   868.00    //freqency in MHz

// Connections to RFM69 module and connected to RH69 chip
#define RFM69_CS    D10       // SPI bus chip select
#define RFM69_INT   D2        // Interrupt pin
#define RFM69_RST   D8        // Reset pin

// Buttons
#define PW_BUTA     A0        // Main user interaction button - used with hardware reset to disable radio
#define PW_BUTB     B1        // Secondary button, also linked to main ESP32 bootloader - not used yet

// LEDs on the main Project Week board
#define RED_LED     A3        // Red 5mm LED - used for radio lock-out indicator
#define GRN_LED     A2        // Green 5mm LED - used for radio connection and packet arrival
#define BI_LEDA     A7        // Bi-colour 5mm LED (A) - not used yet
#define BI_LEDB     D3        // Bi-colour 5mm LED (B) - not used yet

// PWM values for the green LED, saves a little power rather than drive at full brightness
#define LED_LOW     10        // Low brightness PWM value 0-255
#define LED_HIGH    60       // High brightness PWM value 0-255

// OLED matrix display module
#define SMOL_OLED             // Comment out if using larger OLED 128x32 instead of smol 96x16
#ifndef SMOL_OLED
#define SCREEN_WIDTH    128   // OLED display width, in pixels
#define SCREEN_HEIGHT   32    // OLED display height, in pixels
#else
#define SCREEN_WIDTH    96   // OLED display width, in pixels
#define SCREEN_HEIGHT   16    // OLED display height, in pixels
#endif
#define SCREEN_ADDRESS  0x3C  // I2C address


// Create a global rfm69hcw object for the radio module
RH_RF69 rf69(RFM69_CS, RFM69_INT);

// Creat a global display object for the OLED module
Adafruit_SSD1306 display(SCREEN_WIDTH, SCREEN_HEIGHT, &Wire);


void setup() 
{
  // Initialize the OLED display
  if (!display.begin(SSD1306_SWITCHCAPVCC, 0x3C)) {
    //Serial.println(F("SSD1306 allocation failed"));
    for (;;); // Don't proceed, loop forever
  }
  display.clearDisplay();
  display.display();

  // Set up LED pins and buttons
  pinMode(RED_LED, OUTPUT);
  pinMode(GRN_LED, OUTPUT);
  pinMode(BI_LEDA, OUTPUT);
  pinMode(BI_LEDB, OUTPUT);
  pinMode(PW_BUTA, INPUT_PULLUP);
  pinMode(RFM69_RST, OUTPUT);


  // Ensure LEDs are off at first
  digitalWrite(GRN_LED, LOW);
  digitalWrite(RED_LED, LOW);
  digitalWrite(BI_LEDA, LOW);
  digitalWrite(BI_LEDB, LOW);
  
  // Pull down radio reset
  digitalWrite(RFM69_RST, LOW);

  // If button pressed when uC is reset prevent radio startup
  // also indicated by flashing RED LED and message on OLED display
  if(digitalRead(PW_BUTA) == LOW) {

    display.clearDisplay();                         
    display.setCursor(0, SCREEN_HEIGHT/2);
    display.setTextSize(1);               
    display.setTextColor(SSD1306_WHITE); 
    display.println("Radio DISABLED");
    display.display();
    
    // infinite loop to flash RED LED
    for(;;){
      digitalWrite(RED_LED, HIGH);   
      delay(250);
      digitalWrite(RED_LED, LOW);   
      delay(250);
    }
  
  }

  // If code makes here, we're good to go. 

  // Open the serial port - useful if we need it for debug and or to Tx/Rx data over the UART
  Serial.begin(115200);
  delay(1000);

 
  // Let's try and start the radio module and light up the GREEN LED if successful
  digitalWrite(GRN_LED, LOW);
  if( rf69.init() ){
    analogWrite(GRN_LED, LED_LOW); // use PWM output to dimly LED
  }

  // Set the frequency - ideally we should check return status to trap for out-of-bounds requests
  rf69.setFrequency(RF69_FREQ);
  
  // Set the transmit power - even if we're receiving we need to set this
  rf69.setTxPower(-2, true);
  
  // Set the receiver configuration for Gaussian FSK, 250 kbits per second, 250kHz bandwidth
  rf69.setModemConfig(RH_RF69::GFSK_Rb250Fd250);

  // The encryption key has to be the same as the one in the transmitter
  uint8_t key[] = { 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08,
                    0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08};
  rf69.setEncryptionKey(key);
  
  // Dump out the register values - useful for debugging what's going on
  //rf69.printRegisters();

}

void display_rssi(int rssi_val)
{

  int barwidth = map(rssi_val, -100, -20, 0, SCREEN_WIDTH);
  display.clearDisplay(); // Clear the display buffer

  // Draw a bargraph of RSSI
#ifndef SMOL_OLED
  // Larger 128x32 OLED display
  display.fillRect(0, 0, barwidth, 10, SSD1306_WHITE);     // Draw filled rectangle
  display.drawRect(0, 0, SCREEN_WIDTH, 10, SSD1306_WHITE); // Outline box
  display.setCursor(0, 16);                                // Set text cursor position
  display.setTextSize(1);                                  // Set text size
#else
  // Smaller 96x16 OLED display
  display.fillRect(0, 0, barwidth, 6, SSD1306_WHITE);     // Draw filled rectangle
  display.drawRect(0, 0, SCREEN_WIDTH, 6, SSD1306_WHITE); // Outline box
  display.setCursor(0, 9);                                // Set text cursor position
  display.setTextSize(1);
#endif

  // Display the numerical value
  display.setTextColor(SSD1306_WHITE); // Set text color
  display.print("RSSI: ");
  display.print(rssi_val); // Print value
  display.println(" dBm");

  // Update display
  display.display();
}


void loop_receiver()
{
  // The RSSI (received signal strength indicator) - roughly the power level in dBm at the module input
  int rssi_val;

  // If there is data available at the receiver input ...
  if (rf69.available())
  {
    // Should be a message for us now   
    uint8_t buf[RH_RF69_MAX_MESSAGE_LEN];
    uint8_t len = sizeof(buf);
    if (rf69.recv(buf, &len)) {

      // We've received valid data in the buffer - brighten the green LED to signify the fact
      analogWrite(GRN_LED, LED_HIGH);
    
      //RH_RF69::printBuffer("request: ", buf, len);
      //Serial.print("got request: ");
      //Serial.println((char*)buf);
      
      // Measure the RSSI value ...
      rssi_val = rf69.lastRssi();

      Serial.println(rssi_val);

      // Update the OLED display
      display_rssi(rssi_val);

      // Short pause of 50ms before dimming LED brightness back down
      delay(50); 
      analogWrite(GRN_LED, LED_LOW); // PWM value
            
      // Code to transmit back a reply
      //uint8_t data[] = "And hello back to you";
      //rf69.send(data, sizeof(data));
      //rf69.waitPacketSent();
      //Serial.println("Sent a reply");
    } 
    else {
      // If we get here it means the message is invalid

      // Briefly flash the bicolour LED red to signify an error
      digitalWrite(BI_LEDB, HIGH);
      delay(50);
      digitalWrite(BI_LEDB, LOW);
      
      //Serial.println("recv failed");
    }
  }
}

void loop_transmitter()
{
  uint8_t data[] = "Welcome to Project Week 2";
  rf69.send(data, sizeof(data));
  rf69.waitPacketSent();
  delay(20);
}

void loop()
{
  loop_receiver();
  //loop_transmitter();
}