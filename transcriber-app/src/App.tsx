import React, { useCallback, useEffect, useRef, useState } from 'react';
import MicIcon from './components/icons/MicIcon';
import StopIcon from './components/icons/StopIcon';
import DownloadIcon from './components/icons/DownloadIcon';
import UploadIcon from './components/icons/UploadIcon';
import TrashIcon from './components/icons/TrashIcon';
import PencilIcon from './components/icons/PencilIcon';
import CheckIcon from './components/icons/CheckIcon';
import CancelIcon from './components/icons/CancelIcon';
import type { TranscriptEntry } from './types';

type TranscriptionMode = 'live' | 'file';

declare global {
  interface Window {
    docx: any;
    saveAs: (blob: Blob, filename: string) => void;
  }
}

const TRANSCRIBE_ENDPOINT = (import.meta.env.VITE_TRANSCRIBE_ENDPOINT as string | undefined) || '/api/transcribe';

const speakerColors = ['#60a5fa', '#f87171', '#4ade80', '#c084fc', '#fb923c'];

const getSpeakerColor = (speaker: string): string => {
  const speakerNumMatch = speaker.match(/\d+/);
  if (!speakerNumMatch) return '#9ca3af';
  const speakerNum = parseInt(speakerNumMatch[0], 10);
  return speakerColors[(speakerNum - 1) % speakerColors.length];
};

const App: React.FC = () => {
  const [isRecording, setIsRecording] = useState(false);
  const [status, setStatus] = useState('Готов к транскрипции. Нажмите "Начать запись".');
  const [transcript, setTranscript] = useState<TranscriptEntry[]>([]);
  const [currentTranscription, setCurrentTranscription] = useState('');

  const [activeTab, setActiveTab] = useState<TranscriptionMode>('live');
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [fileUrl, setFileUrl] = useState('');
  const [isFileProcessing, setIsFileProcessing] = useState(false);

  const [editingIndex, setEditingIndex] = useState<number | null>(null);
  const [editingText, setEditingText] = useState('');
  const [editingSpeaker, setEditingSpeaker] = useState('');

  const recorderRef = useRef<MediaRecorder | null>(null);
  const micStreamRef = useRef<MediaStream | null>(null);
  const recordedChunksRef = useRef<Blob[]>([]);

  const docxReady = Boolean(window.docx?.Packer && window.docx?.Document);
  const fileSaverReady = typeof window.saveAs === 'function';

  useEffect(() => {
    if (!fileSaverReady) {
      console.error('FileSaver.js (saveAs) not found. Text and DOCX downloads will be disabled.');
    }
    if (!docxReady) {
      console.error('docx library not found. DOCX downloads will be disabled.');
    }
  }, [docxReady, fileSaverReady]);

  const handleStopRecording = useCallback(() => {
    if (!isRecording) return;

    setIsRecording(false);
    setStatus('Обработка финальной транскрипции...');

    if (recorderRef.current && recorderRef.current.state !== 'inactive') {
      recorderRef.current.stop();
    }

    if (micStreamRef.current) {
      micStreamRef.current.getTracks().forEach((track) => track.stop());
      micStreamRef.current = null;
    }
  }, [isRecording]);

  const sendBlobForTranscription = useCallback(
    async (audioBlob: Blob, source: 'live' | 'file', fileName?: string) => {
      const formData = new FormData();
      formData.append('audio', audioBlob, fileName || `${source}-recording.webm`);

      setStatus('Отправка на серверную транскрипцию...');

      try {
        const response = await fetch(TRANSCRIBE_ENDPOINT, {
          method: 'POST',
          body: formData,
        });

        if (!response.ok) {
          const message = await response.text();
          throw new Error(message || 'Сервер вернул ошибку');
        }

        const data = (await response.json()) as { transcript?: string };
        const text = data.transcript?.trim();

        if (text) {
          setTranscript([
            {
              speaker: 'Спикер',
              text,
              timestamp: new Date().toLocaleTimeString('ru-RU'),
            },
          ]);
          setStatus('Транскрипция завершена.');
          setCurrentTranscription('');
        } else {
          setStatus('Транскрипция завершена, но текст пуст.');
        }
      } catch (error) {
        console.error('Ошибка транскрипции:', error);
        const message = error instanceof Error ? error.message : String(error);
        setStatus(`Ошибка: ${message}`);
      }
    },
    [],
  );

  const handleStartRecording = useCallback(async () => {
    if (isRecording) return;
    setTranscript([]);
    setCurrentTranscription('');
    setStatus('Инициализация...');

    try {
      const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
      micStreamRef.current = stream;

      const recorder = new MediaRecorder(stream, { mimeType: 'audio/webm' });
      recorderRef.current = recorder;
      recordedChunksRef.current = [];

      recorder.ondataavailable = (event) => {
        if (event.data.size > 0) {
          recordedChunksRef.current.push(event.data);
        }
      };

      recorder.onstop = async () => {
        const blob = new Blob(recordedChunksRef.current, { type: 'audio/webm' });
        recordedChunksRef.current = [];
        await sendBlobForTranscription(blob, 'live');
      };

      recorder.onerror = (event) => {
        console.error('Ошибка записи аудио:', event.error);
        setStatus(`Ошибка записи аудио: ${event.error?.message ?? 'неизвестная ошибка'}`);
      };

      recorder.start();
      setStatus('Запись началась. Говорите в микрофон.');
      setCurrentTranscription('Запись активна...');
      setIsRecording(true);
    } catch (error) {
      console.error('Ошибка начала записи:', error);
      setStatus('Ошибка: Не удалось получить доступ к микрофону. Проверьте разрешения.');
      if (micStreamRef.current) {
        micStreamRef.current.getTracks().forEach((track) => track.stop());
        micStreamRef.current = null;
      }
    }
  }, [isRecording, sendBlobForTranscription]);

  const handleClearTranscript = () => {
    setTranscript([]);
    setCurrentTranscription('');
    setStatus('Транскрипция очищена.');
  };

  const handleDownloadDocx = () => {
    if (!transcript.length) {
      alert('Нет транскрипции для скачивания.');
      return;
    }
    if (!docxReady || !fileSaverReady) {
      alert('Необходимые для скачивания библиотеки не загружены.');
      return;
    }

    const { Paragraph, TextRun, AlignmentType, Document, Packer } = window.docx;

    const paragraphs = [
      new Paragraph({
        children: [new TextRun({ text: 'Аудио транскрипция', bold: true, size: 32 })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 },
      }),
      new Paragraph({
        children: [new TextRun({ text: `Сгенерировано: ${new Date().toLocaleString('ru-RU')}`, italics: true, size: 20 })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 600 },
      }),
    ];

    const hasSpeakerTags = transcript.some((entry) => /Спикер\s\d+:?/.test(entry.speaker));

    transcript.forEach((entry) => {
      if (!entry.text.trim()) return;

      const speakerLineChildren: any[] = [];
      if (entry.timestamp) {
        speakerLineChildren.push(new TextRun({ text: `[${entry.timestamp}] `, bold: true, color: '555555' }));
      }
      if (hasSpeakerTags) {
        speakerLineChildren.push(
          new TextRun({ text: `${entry.speaker}:`, bold: true, color: getSpeakerColor(entry.speaker).slice(1) }),
        );
      }

      if (speakerLineChildren.length) {
        paragraphs.push(
          new Paragraph({
            children: speakerLineChildren,
            spacing: { after: 100 },
          }),
        );
      }

      paragraphs.push(
        new Paragraph({
          children: [new TextRun({ text: entry.text.trim() })],
          spacing: { after: 400 },
          indent: { left: hasSpeakerTags ? 720 : 0 },
        }),
      );
    });

    const doc = new Document({ sections: [{ children: paragraphs }] });

    Packer.toBlob(doc).then((blob: Blob) => {
      window.saveAs(blob, 'transcript.docx');
    });
  };

  const handleDownloadTxt = () => {
    if (!transcript.length) {
      alert('Нет транскрипции для скачивания.');
      return;
    }
    if (!fileSaverReady) {
      alert('Библиотека для скачивания не загружена.');
      return;
    }

    const hasSpeakerTags = transcript.some((entry) => /Спикер\s\d+:?/.test(entry.speaker));

    const content = [
      'Аудио транскрипция',
      `Сгенерировано: ${new Date().toLocaleString('ru-RU')}`,
      '',
      ...transcript.map((entry) => {
        const parts: string[] = [];
        if (entry.timestamp) parts.push(`[${entry.timestamp}]`);
        if (hasSpeakerTags) parts.push(`${entry.speaker}:`);
        parts.push(entry.text.trim());
        return parts.join(' ');
      }),
    ].join('\n\n');

    const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
    window.saveAs(blob, 'transcript.txt');
  };

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    if (event.target.files?.[0]) {
      setSelectedFile(event.target.files[0]);
      setFileUrl('');
    }
  };

  const handleTranscribeFile = async () => {
    if (!selectedFile && !fileUrl) {
      setStatus('Пожалуйста, выберите файл или укажите URL.');
      return;
    }

    setTranscript([]);
    setIsFileProcessing(true);
    setStatus('Подготовка файла...');

    try {
      let audioBlob: Blob;
      let mimeType: string;

      if (selectedFile) {
        audioBlob = selectedFile;
        mimeType = selectedFile.type || 'audio/mpeg';
      } else {
        setStatus('Загрузка файла по URL...');
        const response = await fetch(fileUrl);
        if (!response.ok) throw new Error(`Не удалось загрузить файл: ${response.statusText}`);
        const arrayBuffer = await response.arrayBuffer();
        mimeType = response.headers.get('Content-Type') || 'audio/mpeg';
        audioBlob = new Blob([arrayBuffer], { type: mimeType });
      }

      await sendBlobForTranscription(audioBlob, 'file', selectedFile?.name || 'file-upload.webm');
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      setStatus(`Ошибка: ${message}`);
      console.error('Ошибка транскрипции файла:', error);
    } finally {
      setIsFileProcessing(false);
    }
  };

  const handleEditStart = (index: number) => {
    setEditingIndex(index);
    setEditingText(transcript[index].text);
    setEditingSpeaker(transcript[index].speaker);
  };

  const handleEditCancel = () => {
    setEditingIndex(null);
    setEditingText('');
    setEditingSpeaker('');
  };

  const handleEditSave = () => {
    if (editingIndex === null) return;
    setTranscript((prev) => {
      const updated = [...prev];
      updated[editingIndex] = {
        ...updated[editingIndex],
        text: editingText,
        speaker: editingSpeaker,
      };
      return updated;
    });
    handleEditCancel();
  };

  return (
    <div className="min-h-screen bg-gray-900 text-gray-100 flex flex-col p-4 md:p-8 font-sans">
      <header className="text-center mb-8">
        <h1 className="text-4xl md:text-5xl font-bold text-transparent bg-clip-text bg-gradient-to-r from-blue-400 to-purple-500">
          Аудио транскрибатор Gemini
        </h1>
        <p className="text-gray-400 mt-2">Транскрипция в реальном времени и из файлов с экспортом в DOCX.</p>
      </header>

      <main className="flex-grow flex flex-col md:flex-row gap-8">
        <div className="w-full md:w-1/3 flex flex-col gap-6">
          <div className="bg-gray-800 p-6 rounded-2xl shadow-lg">
            <div className="flex border-b border-gray-700 mb-4">
              <button
                type="button"
                onClick={() => setActiveTab('live')}
                className={`py-2 px-4 font-semibold transition-colors duration-200 ${
                  activeTab === 'live' ? 'border-b-2 border-purple-400 text-purple-400' : 'text-gray-400 hover:text-white'
                }`}
              >
                В реальном времени
              </button>
              <button
                type="button"
                onClick={() => setActiveTab('file')}
                className={`py-2 px-4 font-semibold transition-colors duration-200 ${
                  activeTab === 'file' ? 'border-b-2 border-purple-400 text-purple-400' : 'text-gray-400 hover:text-white'
                }`}
              >
                Из файла
              </button>
            </div>

            {activeTab === 'live' ? (
              <div className="flex flex-col items-center gap-4">
                <button
                  type="button"
                  onClick={isRecording ? () => handleStopRecording() : () => handleStartRecording()}
                  disabled={isFileProcessing}
                  className={`flex items-center justify-center gap-2 w-48 px-6 py-3 text-white font-semibold rounded-full shadow-md transition-all duration-200 disabled:bg-gray-600 disabled:cursor-not-allowed ${
                    isRecording ? 'bg-red-600 hover:bg-red-700 focus:ring-red-500' : 'bg-blue-600 hover:bg-blue-700 focus:ring-blue-500'
                  }`}
                >
                  {isRecording ? <StopIcon className="w-6 h-6" /> : <MicIcon className="w-6 h-6" />}
                  <span>{isRecording ? 'Остановить' : 'Начать запись'}</span>
                </button>
                <p className="text-sm text-gray-400 text-center h-10">{status}</p>
              </div>
            ) : (
              <div className="flex flex-col items-center gap-4">
                <label className="w-full flex items-center justify-center gap-2 px-6 py-3 bg-gray-700 text-blue-400 font-semibold rounded-full shadow-md hover:bg-gray-600 cursor-pointer transition-colors duration-200">
                  <UploadIcon className="w-6 h-6" />
                  <span className="truncate">{selectedFile ? selectedFile.name : 'Выбрать файл'}</span>
                  <input type="file" className="hidden" onChange={handleFileChange} accept="audio/*" disabled={isFileProcessing} />
                </label>
                <div className="text-sm text-gray-500">или</div>
                <input
                  type="text"
                  value={fileUrl}
                  onChange={(event) => {
                    setFileUrl(event.target.value);
                    setSelectedFile(null);
                  }}
                  placeholder="Вставьте URL аудиофайла"
                  className="w-full bg-gray-900 text-gray-200 px-4 py-2 rounded-full border border-gray-700 focus:outline-none focus:ring-2 focus:ring-purple-500"
                  disabled={isFileProcessing}
                />

                <button
                  type="button"
                  onClick={handleTranscribeFile}
                  disabled={isFileProcessing || isRecording || (!selectedFile && !fileUrl)}
                  className="w-full px-6 py-3 mt-2 bg-purple-600 text-white font-semibold rounded-full shadow-md hover:bg-purple-700 focus:outline-none focus:ring-2 focus:ring-purple-500 transition-all duration-200 disabled:bg-gray-600 disabled:cursor-not-allowed"
                >
                  <span>{isFileProcessing ? 'Обработка...' : 'Транскрибировать файл'}</span>
                </button>

                <p className="text-sm text-gray-400 text-center h-5 mt-2">{status}</p>
              </div>
            )}
          </div>

          <div className={`bg-gray-800 p-6 rounded-2xl shadow-lg flex-grow flex flex-col ${activeTab !== 'live' ? 'hidden' : ''}`}>
            <h2 className="text-lg font-semibold text-purple-400 mb-3 flex items-center gap-2">
              <div className={`w-3 h-3 rounded-full ${isRecording ? 'bg-red-500 animate-pulse' : 'bg-gray-500'}`} />
              Текущая транскрипция
            </h2>
            <div className="bg-gray-900 rounded-lg p-4 h-32 flex-grow overflow-y-auto text-gray-300">
              {currentTranscription || <span className="text-gray-500">...</span>}
            </div>
          </div>
        </div>

        <div className="w-full md:w-2/3 bg-gray-800 p-6 rounded-2xl shadow-lg flex flex-col">
          <div className="flex justify-between items-center mb-4 flex-wrap gap-2">
            <h2 className="text-2xl font-bold text-purple-400">Транскрипция</h2>
            <div className="flex items-center gap-2 flex-wrap">
              <button
                type="button"
                onClick={handleClearTranscript}
                disabled={isRecording || isFileProcessing || transcript.length === 0}
                className="flex items-center gap-2 px-4 py-2 bg-amber-600 text-white font-semibold rounded-full shadow-md hover:bg-amber-700 disabled:bg-gray-600 disabled:cursor-not-allowed"
              >
                <TrashIcon className="w-5 h-5" />
                <span>Очистить</span>
              </button>
              <button
                type="button"
                onClick={handleDownloadTxt}
                disabled={!fileSaverReady || transcript.length === 0}
                title={!fileSaverReady ? 'Библиотека не загружена' : 'Скачать TXT'}
                className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white font-semibold rounded-full shadow-md hover:bg-blue-700 disabled:bg-gray-600 disabled:cursor-not-allowed"
              >
                <DownloadIcon className="w-5 h-5" />
                <span>TXT</span>
              </button>
              <button
                type="button"
                onClick={handleDownloadDocx}
                disabled={!docxReady || transcript.length === 0}
                title={!docxReady ? 'Библиотека не загружена' : 'Скачать DOCX'}
                className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white font-semibold rounded-full shadow-md hover:bg-blue-700 disabled:bg-gray-600 disabled:cursor-not-allowed"
              >
                <DownloadIcon className="w-5 h-5" />
                <span>DOCX</span>
              </button>
            </div>
          </div>
          <div className="flex-grow bg-gray-900 rounded-lg p-4 overflow-y-auto space-y-4">
            {transcript.length > 0 ? (
              transcript.map((entry, index) => (
                <div
                  key={`${entry.speaker}-${index}`}
                  className="flex flex-col p-3 rounded-lg bg-gray-800/50 transition-colors duration-200 hover:bg-gray-800/80"
                >
                  {editingIndex === index ? (
                    <div className="flex flex-col gap-2">
                      <div className="flex items-center gap-2">
                        <input
                          type="text"
                          value={editingSpeaker}
                          onChange={(event) => setEditingSpeaker(event.target.value)}
                          className="bg-gray-700 text-sm font-semibold p-1 rounded w-32 focus:outline-none focus:ring-2 focus:ring-purple-500"
                          style={{ color: getSpeakerColor(editingSpeaker) }}
                        />
                        <span className="text-sm text-gray-500">{entry.timestamp ? `- ${entry.timestamp}` : ''}</span>
                        <div className="flex-grow" />
                        <button
                          type="button"
                          onClick={handleEditSave}
                          className="text-green-400 hover:text-green-300"
                        >
                          <CheckIcon className="w-5 h-5" />
                        </button>
                        <button
                          type="button"
                          onClick={handleEditCancel}
                          className="text-red-400 hover:text-red-300"
                        >
                          <CancelIcon className="w-5 h-5" />
                        </button>
                      </div>
                      <textarea
                        value={editingText}
                        onChange={(event) => setEditingText(event.target.value)}
                        className="w-full bg-gray-700 text-gray-200 p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-purple-500"
                        rows={3}
                      />
                    </div>
                  ) : (
                    <>
                      <div className="flex items-center text-sm font-medium text-gray-400 mb-1">
                        {/Спикер\s\d+:?/.test(entry.speaker) ? (
                          <span className="font-semibold" style={{ color: getSpeakerColor(entry.speaker) }}>
                            {entry.speaker}
                          </span>
                        ) : (
                          <span className="font-semibold text-gray-400">{entry.speaker}</span>
                        )}
                        {entry.timestamp ? ` - ${entry.timestamp}` : ''}
                        <div className="flex-grow" />
                        <button
                          type="button"
                          onClick={() => handleEditStart(index)}
                          disabled={editingIndex !== null}
                          className="text-gray-500 hover:text-blue-400 disabled:opacity-50 disabled:cursor-not-allowed"
                        >
                          <PencilIcon className="w-4 h-4" />
                        </button>
                      </div>
                      <p className="text-gray-200 ml-1 whitespace-pre-line">{entry.text}</p>
                    </>
                  )}
                </div>
              ))
            ) : (
              <div className="flex items-center justify-center h-full text-gray-500">
                Ваша транскрипция появится здесь...
              </div>
            )}
          </div>
        </div>
      </main>
    </div>
  );
};

export default App;
