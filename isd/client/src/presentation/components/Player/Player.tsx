import { useAudioPlayer } from '../../hooks/useAudioPlayer';
import { Play, Pause, SkipBack, SkipForward, Volume2, Repeat, Loader2 } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Slider } from '@/components/ui/slider';
import { cn } from '@/lib/utils';

export const Player = () => {
  const {
    currentTrack,
    playerState,
    currentTime,
    duration,
    volume,
    isLooping,
    pause,
    resume,
    seek,
    setVolume,
    toggleLoop,
  } = useAudioPlayer();

  const formatTime = (seconds: number): string => {
    if (!isFinite(seconds)) return '0:00';
    const mins = Math.floor(seconds / 60);
    const secs = Math.floor(seconds % 60);
    return `${mins}:${secs.toString().padStart(2, '0')}`;
  };

  const handlePlayPause = () => {
    if (playerState === 'playing') {
      pause();
    } else if (playerState === 'paused') {
      resume();
    }
  };

  const handleSeek = (value: number[]) => {
    seek(value[0]);
  };

  const handleVolumeChange = (value: number[]) => {
    setVolume(value[0]);
  };

  if (!currentTrack) {
    return (
      <div className="fixed bottom-0 left-0 right-0 bg-card border-t border-border p-4">
        <div className="max-w-7xl mx-auto text-center text-muted-foreground">
          No track selected
        </div>
      </div>
    );
  }

  const isLoading = playerState === 'loading' || playerState === 'buffering';
  const isPlayable = playerState === 'playing' || playerState === 'paused';

  return (
    <div className="fixed bottom-0 left-0 right-0 bg-card border-t border-border shadow-lg">
      <div className="max-w-7xl mx-auto p-4">
        {/* Progress Bar */}
        <div className="mb-4">
          <Slider
            value={[currentTime]}
            max={duration || 100}
            step={0.1}
            onValueChange={handleSeek}
            disabled={!isPlayable}
            className="w-full"
          />
          <div className="flex justify-between text-xs text-muted-foreground mt-1">
            <span>{formatTime(currentTime)}</span>
            <span>{formatTime(duration)}</span>
          </div>
        </div>

        <div className="flex items-center justify-between gap-4">
          {/* Track Info */}
          <div className="flex items-center gap-3 flex-1 min-w-0">
            {currentTrack.coverUrl && (
              <img
                src={currentTrack.coverUrl}
                alt={currentTrack.title}
                className="w-12 h-12 rounded object-cover"
              />
            )}
            <div className="min-w-0 flex-1">
              <div className="font-medium truncate">{currentTrack.title}</div>
              <div className="text-sm text-muted-foreground truncate">
                {currentTrack.artist}
              </div>
            </div>
          </div>

          {/* Controls */}
          <div className="flex items-center gap-2">
            <Button variant="ghost" size="icon" disabled>
              <SkipBack className="h-5 w-5" />
            </Button>
            
            <Button
              variant="default"
              size="icon"
              onClick={handlePlayPause}
              disabled={isLoading}
              className="h-10 w-10"
            >
              {isLoading ? (
                <Loader2 className="h-5 w-5 animate-spin" />
              ) : playerState === 'playing' ? (
                <Pause className="h-5 w-5" />
              ) : (
                <Play className="h-5 w-5" />
              )}
            </Button>

            <Button variant="ghost" size="icon" disabled>
              <SkipForward className="h-5 w-5" />
            </Button>

            <Button
              variant="ghost"
              size="icon"
              onClick={toggleLoop}
              className={cn(isLooping && 'text-primary')}
            >
              <Repeat className="h-5 w-5" />
            </Button>
          </div>

          {/* Volume */}
          <div className="flex items-center gap-2 flex-1 justify-end">
            <Volume2 className="h-5 w-5 text-muted-foreground" />
            <Slider
              value={[volume]}
              max={1}
              step={0.01}
              onValueChange={handleVolumeChange}
              className="w-24"
            />
          </div>
        </div>
      </div>
    </div>
  );
};
