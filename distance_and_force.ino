
#include <SoftwareSerial.h>

class SimpleKalmanFilter {
private:
  float Q, R, P, K, x;
public:
  SimpleKalmanFilter(float q, float r, float p) : Q(q), R(r), P(p), K(0), x(0) {}
  float update(float z) {
    P += Q; K = P / (P + R); x += K * (z - x); P = (1 - K) * P; return x;
  }
};

// --- 引脚定义 ---
#define IR_RX_PIN 9
#define IR_TX_PIN 10
#define FORCE_SENSOR_PIN A1
#define PRESSURE_SENSOR_PIN A0

SoftwareSerial irSensorSerial(IR_RX_PIN, IR_TX_PIN);

// --- 传感器变量 ---
SimpleKalmanFilter distanceKalman(0.01, 0.1, 0.5);
SimpleKalmanFilter forceKalman(0.01, 0.1, 0.5);
const int HISTORY_SIZE = 20;
float distanceHistory[HISTORY_SIZE], forceHistory[HISTORY_SIZE];
int distanceHistoryIndex = 0, forceHistoryIndex = 0;
bool bufferIsFull = false;
float stable_distance_mm = -1.0, stable_force_N = 0.0;

unsigned long lastSendTime = 0;
const unsigned long SEND_INTERVAL = 100;

void setup() {
  Serial.begin(9600);
  irSensorSerial.begin(115200);
  pinMode(FORCE_SENSOR_PIN, INPUT);
  pinMode(PRESSURE_SENSOR_PIN, INPUT);
  
  // !!! 这是最重要的修改 !!!
  // 在所有初始化完成后，向Python发送一个清晰的“准备就绪”信号。
  Serial.println("Arduino is Ready"); 
}

void loop() {
  processDistanceSensor(); 
  processForceSensor();
  
  if (millis() - lastSendTime >= SEND_INTERVAL) {
    lastSendTime = millis();
    sendAllData();
  }
}

// ... 所有其他函数 (processPressureSensor, sendAllData, processForceSensor, 等) ...
// ... 与我们之前最终确认的版本完全相同，这里为简洁省略 ...
// ... 请确保你使用的是包含所有这些函数的完整版本 ...
float processPressureSensor() {
  int sensorValue = analogRead(PRESSURE_SENSOR_PIN);
  float pressure = mapFloat(sensorValue, 0.0, 1023.0, 0.0, 1.0);
  return pressure;
}

void sendAllData() {
  if (!bufferIsFull) {
    if (distanceHistoryIndex >= HISTORY_SIZE - 1 || forceHistoryIndex >= HISTORY_SIZE - 1) {
      bufferIsFull = true;
    }
    return;
  }
  float current_pressure = processPressureSensor();
  Serial.print(stable_force_N, 2);
  Serial.print(",");
  Serial.print(stable_distance_mm, 2);
  Serial.print(",");
  Serial.println(current_pressure, 3);
}

void processForceSensor() {
  float rawForce = mapFloat(analogRead(FORCE_SENSOR_PIN), 0.0, 1023.0, 0.0, 200.0);
  float kalmanForce = forceKalman.update(rawForce);
  forceHistory[forceHistoryIndex] = kalmanForce;
  float sum_force = 0.0;
  for(int i=0; i<HISTORY_SIZE; ++i) sum_force += forceHistory[i];
  stable_force_N = sum_force / HISTORY_SIZE;
  forceHistoryIndex = (forceHistoryIndex + 1) % HISTORY_SIZE;
}

void processDistanceSensor() {
  static enum { WAITING_FOR_AC, WAITING_FOR_CA, READING_PAYLOAD } state = WAITING_FOR_AC;
  static byte payloadBuffer[8];
  static int bufferIndex = 0;
  if (irSensorSerial.available() > 0) {
    byte b = irSensorSerial.read();
    switch (state) {
      case WAITING_FOR_AC: if (b == 0xAC) state = WAITING_FOR_CA; break;
      case WAITING_FOR_CA: state = (b == 0xCA) ? READING_PAYLOAD : WAITING_FOR_AC; bufferIndex = 0; break;
      case READING_PAYLOAD:
        payloadBuffer[bufferIndex++] = b;
        if (bufferIndex >= 8) {
          if (payloadBuffer[5] != 0x00 && payloadBuffer[6] == 0xDC && payloadBuffer[7] == 0xCD) {
            long raw_dist = ((long)payloadBuffer[1] << 24)|((long)payloadBuffer[2] << 16)|((long)payloadBuffer[3] << 8)|(long)payloadBuffer[4];
            float kalmanDist = distanceKalman.update((float)raw_dist / 100.0f);
            distanceHistory[distanceHistoryIndex] = kalmanDist;
            float sum_dist = 0.0;
            for(int i=0; i<HISTORY_SIZE; ++i) sum_dist += distanceHistory[i];
            stable_distance_mm = sum_dist / HISTORY_SIZE;
            distanceHistoryIndex = (distanceHistoryIndex + 1) % HISTORY_SIZE;
          }
          state = WAITING_FOR_AC;
        }
        break;
    }
  }
}

float mapFloat(float x, float in_min, float in_max, float out_min, float out_max) {
  return (x - in_min) * (out_max - out_min) / (in_max - in_min) + out_min;
}