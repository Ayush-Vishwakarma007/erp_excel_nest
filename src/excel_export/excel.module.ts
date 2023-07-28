/* eslint-disable prettier/prettier */
import { Module } from '@nestjs/common';
import { MulterModule } from '@nestjs/platform-express';
import { ExcelController } from './excel.controller';

@Module({
    imports: [
        MulterModule.register({
            dest: 'uploads',
        }),
    ],
    controllers: [ExcelController],
})
export class ExcelModule { }

