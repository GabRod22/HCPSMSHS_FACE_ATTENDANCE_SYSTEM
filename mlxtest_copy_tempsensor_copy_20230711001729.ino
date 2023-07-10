// Import the Adafruit_MLX90614 library
#include <Adafruit_MLX90614.h>

// The temperature sensor
Adafruit_MLX90614 mlx = Adafruit_MLX90614();

void setup() {
  // Initialize the temperature sensor
  mlx.begin();

  // Initialize the serial monitor
  Serial.begin(9600);
}

void loop() {
  // Read the temperature from the sensor
  float temperature = mlx.readObjectTempC();

  // Print the temperature to the serial monitor
  Serial.print(temperature);
  Serial.println("C");

  // Delay for 500 milliseconds
  delay(500);
}
