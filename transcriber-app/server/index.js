import express from 'express';
import cors from 'cors';
import multer from 'multer';
import { GoogleGenAI } from '@google/genai';

const app = express();
const upload = multer();

app.use(cors());

const apiKey = process.env.GEMINI_API_KEY || process.env.GOOGLE_API_KEY;
let client = null;

if (apiKey) {
  client = new GoogleGenAI({ apiKey });
} else {
  console.warn('[transcriber-app] GEMINI_API_KEY is not set. Transcription requests will fail.');
}

app.post('/api/transcribe', upload.single('audio'), async (req, res) => {
  if (!client) {
    res.status(500).send('Серверная конфигурация отсутствует: не указан GEMINI_API_KEY.');
    return;
  }

  if (!req.file) {
    res.status(400).send('Не удалось получить аудиофайл.');
    return;
  }

  try {
    const base64Data = req.file.buffer.toString('base64');
    const mimeType = req.file.mimetype || 'audio/webm';

    const response = await client.models.generateContent({
      model: 'gemini-2.5-pro',
      contents: {
        parts: [
          {
            text:
              "Транскрибируй это аудио. Если в аудио несколько говорящих, раздели их реплики, начиная каждую с новой строки в формате 'Спикер N: [текст реплики]'. Если говорящий один, просто предоставь сплошной текст транскрипции.",
          },
          { inlineData: { mimeType, data: base64Data } },
        ],
      },
    });

    res.json({ transcript: response.text ?? '' });
  } catch (error) {
    console.error('Ошибка сервера транскрипции:', error);
    res.status(500).send('Не удалось выполнить транскрипцию.');
  }
});

const port = process.env.PORT || 8788;
app.listen(port, () => {
  console.log(`[transcriber-app] Proxy сервер запущен на порту ${port}`);
});
