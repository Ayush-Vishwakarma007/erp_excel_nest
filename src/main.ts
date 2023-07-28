import { NestFactory } from '@nestjs/core';
import { AppModule } from './app.module';

async function bootstrap() {
  const app = await NestFactory.create(AppModule, { cors: true });
  const ipAddress = '192.168.21.10';
  const port = 5000;
  await app.listen(port, ipAddress, () => {
    console.log(`App is running on http://${ipAddress}:${port}`);
  });
}
bootstrap();
