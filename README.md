# 6 7 Motion Counter

A simple browser-based motion counter that uses a webcam and pose tracking to count a basic "6 7" arm motion.

The app uses MediaPipe Pose Landmarker to detect body landmarks from the camera feed. It tracks the position of the wrists and shoulders, then counts one repetition when the user raises the right hand followed by the left hand.

## Demo

This project is meant to run directly in the browser.

If GitHub Pages is enabled, the site should be available at:

```text
https://adamyellowtiger.github.io/email-newsletter/
