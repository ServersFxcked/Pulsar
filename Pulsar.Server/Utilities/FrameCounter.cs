using System;
using System.Collections.Generic;
using System.Linq;

namespace Pulsar.Server.Utilities
{
    public class FrameUpdatedEventArgs : EventArgs
    {
        public float CurrentFramesPerSecond { get; private set; }

        public FrameUpdatedEventArgs(float _CurrentFramesPerSecond)
        {
            CurrentFramesPerSecond = _CurrentFramesPerSecond;
        }
    }

    public delegate void FrameUpdatedEventHandler(FrameUpdatedEventArgs e);

    public class FrameCounter
    {
        public long TotalFrames { get; private set; }
        public float TotalSeconds { get; private set; }
        public float AverageFramesPerSecond { get; private set; }

        public const int MAXIMUM_SAMPLES = 100;

        private Queue<float> _sampleBuffer = new Queue<float>();
        private float _currentSampleSum = 0f; // Add this line

        public event FrameUpdatedEventHandler FrameUpdated;

        public void Update(float deltaTime)
        {
            float currentFramesPerSecond = 1.0f / deltaTime;

            _sampleBuffer.Enqueue(currentFramesPerSecond);

            if (_sampleBuffer.Count == MAXIMUM_SAMPLES) // Check if buffer is full before dequeueing
            {
                _currentSampleSum -= _sampleBuffer.Dequeue(); // Subtract the oldest sample from the sum
            }
            _sampleBuffer.Enqueue(currentFramesPerSecond);
            _currentSampleSum += currentFramesPerSecond; // Add the new sample to the sum

            AverageFramesPerSecond = _currentSampleSum / _sampleBuffer.Count; // Calculate average

            OnFrameUpdated(new FrameUpdatedEventArgs(AverageFramesPerSecond));

            TotalFrames++;
            TotalSeconds += deltaTime;
        }

        protected virtual void OnFrameUpdated(FrameUpdatedEventArgs e)
        {
            FrameUpdatedEventHandler handler = FrameUpdated;
            if (handler != null)
                handler(e);
        }
    }
}