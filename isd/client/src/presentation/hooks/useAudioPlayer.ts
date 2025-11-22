import { useState, useEffect, useRef, useCallback } from 'react';
import { Track } from '../../domain/entities/Track';

export type PlayerState = 'idle' | 'loading' | 'buffering' | 'playing' | 'paused' | 'error';

interface UseAudioPlayerReturn {
  currentTrack: Track | null;
  playerState: PlayerState;
  currentTime: number;
  duration: number;
  volume: number;
  isLooping: boolean;
  play: (track: Track) => void;
  pause: () => void;
  resume: () => void;
  stop: () => void;
  seek: (time: number) => void;
  setVolume: (volume: number) => void;
  toggleLoop: () => void;
}

export const useAudioPlayer = (): UseAudioPlayerReturn => {
  const audioRef = useRef<HTMLAudioElement | null>(null);
  const [currentTrack, setCurrentTrack] = useState<Track | null>(null);
  const [playerState, setPlayerState] = useState<PlayerState>('idle');
  const [currentTime, setCurrentTime] = useState(0);
  const [duration, setDuration] = useState(0);
  const [volume, setVolumeState] = useState(0.7);
  const [isLooping, setIsLooping] = useState(false);

  // Initialize audio element
  useEffect(() => {
    const audio = new Audio();
    audio.preload = 'metadata';
    audio.volume = volume;
    audioRef.current = audio;

    return () => {
      audio.pause();
      audio.src = '';
      audioRef.current = null;
    };
  }, []);

  // Setup audio event listeners
  useEffect(() => {
    const audio = audioRef.current;
    if (!audio) return;

    const handleLoadStart = () => {
      setPlayerState('loading');
    };

    const handleCanPlay = () => {
      setPlayerState('playing');
      audio.play().catch((error) => {
        console.error('Playback failed:', error);
        setPlayerState('error');
      });
    };

    const handleWaiting = () => {
      setPlayerState('buffering');
    };

    const handlePlaying = () => {
      setPlayerState('playing');
    };

    const handlePause = () => {
      if (playerState !== 'loading' && playerState !== 'buffering') {
        setPlayerState('paused');
      }
    };

    const handleTimeUpdate = () => {
      setCurrentTime(audio.currentTime);
    };

    const handleDurationChange = () => {
      setDuration(audio.duration);
    };

    const handleEnded = () => {
      if (!isLooping) {
        setPlayerState('idle');
        setCurrentTime(0);
      }
    };

    const handleError = (e: ErrorEvent) => {
      console.error('Audio error:', e);
      setPlayerState('error');
    };

    audio.addEventListener('loadstart', handleLoadStart);
    audio.addEventListener('canplay', handleCanPlay);
    audio.addEventListener('waiting', handleWaiting);
    audio.addEventListener('playing', handlePlaying);
    audio.addEventListener('pause', handlePause);
    audio.addEventListener('timeupdate', handleTimeUpdate);
    audio.addEventListener('durationchange', handleDurationChange);
    audio.addEventListener('ended', handleEnded);
    audio.addEventListener('error', handleError as any);

    return () => {
      audio.removeEventListener('loadstart', handleLoadStart);
      audio.removeEventListener('canplay', handleCanPlay);
      audio.removeEventListener('waiting', handleWaiting);
      audio.removeEventListener('playing', handlePlaying);
      audio.removeEventListener('pause', handlePause);
      audio.removeEventListener('timeupdate', handleTimeUpdate);
      audio.removeEventListener('durationchange', handleDurationChange);
      audio.removeEventListener('ended', handleEnded);
      audio.removeEventListener('error', handleError as any);
    };
  }, [playerState, isLooping]);

  // Update loop attribute
  useEffect(() => {
    if (audioRef.current) {
      audioRef.current.loop = isLooping;
    }
  }, [isLooping]);

  const play = useCallback((track: Track) => {
    const audio = audioRef.current;
    if (!audio) return;

    setCurrentTrack(track);
    setPlayerState('loading');
    audio.src = track.audioUrl;
    audio.load();
  }, []);

  const pause = useCallback(() => {
    const audio = audioRef.current;
    if (!audio) return;
    audio.pause();
  }, []);

  const resume = useCallback(() => {
    const audio = audioRef.current;
    if (!audio) return;
    audio.play().catch((error) => {
      console.error('Resume failed:', error);
      setPlayerState('error');
    });
  }, []);

  const stop = useCallback(() => {
    const audio = audioRef.current;
    if (!audio) return;
    audio.pause();
    audio.currentTime = 0;
    setCurrentTrack(null);
    setPlayerState('idle');
    setCurrentTime(0);
  }, []);

  const seek = useCallback((time: number) => {
    const audio = audioRef.current;
    if (!audio) return;
    audio.currentTime = time;
  }, []);

  const setVolume = useCallback((newVolume: number) => {
    const audio = audioRef.current;
    if (!audio) return;
    const clampedVolume = Math.max(0, Math.min(1, newVolume));
    audio.volume = clampedVolume;
    setVolumeState(clampedVolume);
  }, []);

  const toggleLoop = useCallback(() => {
    setIsLooping((prev) => !prev);
  }, []);

  return {
    currentTrack,
    playerState,
    currentTime,
    duration,
    volume,
    isLooping,
    play,
    pause,
    resume,
    stop,
    seek,
    setVolume,
    toggleLoop,
  };
};
