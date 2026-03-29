from pathlib import Path

def replace_function(text, func_name, new_block):
    target = f"function {func_name}"
    start = text.find(target)
    if start == -1:
        return text
    index = text.find('{', start)
    if index == -1:
        return text
    brace = 0
    i = index
    while i < len(text):
        if text[i] == '{':
            brace += 1
        elif text[i] == '}':
            brace -= 1
            if brace == 0:
                end = i + 1
                break
        i += 1
    else:
        return text
    return text[:start] + new_block + text[end:]

HTML_PATH = Path(__file__).resolve().parents[2] / 'web-assets' / 'rrecision' / 'Rrecision_tilt_sensor.html'
text = HTML_PATH.read_text(encoding='utf-8')

text = text.replace('‚¸“xŒXŽÎŠpƒZƒ“ƒT[', 'ŒXŽÎŠpƒZƒ“ƒT[')
text = text.replace('‚¸“x', '')

css_block = "        .sensor-label {\n            position: absolute;\n            bottom: 12px;\n            right: 22px;\n            font-size: 10px;\n            letter-spacing: 2px;\n            color: rgba(255, 255, 255, 0.4);\n        }\n\n"
if '.sensor-label' not in text:
    text = text.replace('        @keyframes pulse {\n            0%, 100% { opacity: 1; }\n            50% { opacity: 0.6; }\n        }\n\n', css_block + '        @keyframes pulse {\n            0%, 100% { opacity: 1; }\n            50% { opacity: 0.6; }\n        }\n\n')

if '<div class="sensor-label">JJ-07</div>' not in text:
    text = text.replace('\n        </div>\n    </div>\n\n    <script>', '\n            <div class="sensor-label">JJ-07</div>\n\n        </div>\n    </div>\n\n    <script>')

if 'let continuousPanner' not in text:
    text = text.replace('let oscillator = null;\n        let gainNode = null;\n', 'let oscillator = null;\n        let gainNode = null;\n        let continuousPanner = null;\n')

create_tone_block = "function createTone(frequency, volume = 0.1, options = {}) {\n            if (!soundEnabled) return;\n\n            const { type = 'sine', pan = 0 } = options;\n\n            if (continuousMode) {\n                if (!oscillator) {\n                    oscillator = audioContext.createOscillator();\n                    gainNode = audioContext.createGain();\n                    if (audioContext.createStereoPanner) {\n                        continuousPanner = audioContext.createStereoPanner();\n                        gainNode.connect(continuousPanner);\n                        continuousPanner.connect(audioContext.destination);\n                    } else {\n                        continuousPanner = null;\n                        gainNode.connect(audioContext.destination);\n                    }\n                    oscillator.connect(gainNode);\n                    oscillator.start();\n                }\n                oscillator.type = type;\n                oscillator.frequency.setValueAtTime(frequency, audioContext.currentTime);\n                if (continuousPanner) {\n                    continuousPanner.pan.setValueAtTime(Math.max(-1, Math.min(1, pan)), audioContext.currentTime);\n                }\n                gainNode.gain.setValueAtTime(volume, audioContext.currentTime);\n            } else {\n                const now = Date.now();\n                if (now - lastSoundTime < SOUND_INTERVAL) return;\n                lastSoundTime = now;\n\n                const osc = audioContext.createOscillator();\n                const gain = audioContext.createGain();\n                osc.type = type;\n                osc.frequency.value = frequency;\n                gain.gain.value = volume;\n\n                if (audioContext.createStereoPanner) {\n                    const panner = audioContext.createStereoPanner();\n                    panner.pan.value = Math.max(-1, Math.min(1, pan));\n                    osc.connect(gain);\n                    gain.connect(panner);\n                    panner.connect(audioContext.destination);\n                } else {\n                    osc.connect(gain);\n                    gain.connect(audioContext.destination);\n                }\n\n                osc.start();\n                gain.gain.exponentialRampToValueAtTime(0.01, audioContext.currentTime + 0.2);\n                osc.stop(audioContext.currentTime + 0.25);\n            }\n        }\n"
text = replace_function(text, 'createTone', create_tone_block + '\n')

stop_tone_block = "function stopContinuousTone() {\n            if (oscillator && continuousMode) {\n                gainNode.gain.exponentialRampToValueAtTime(0.001, audioContext.currentTime + 0.1);\n                oscillator.stop(audioContext.currentTime + 0.12);\n                oscillator = null;\n                gainNode = null;\n                continuousPanner = null;\n            }\n        }\n"
text = replace_function(text, 'stopContinuousTone', stop_tone_block + '\n')

handle_block = "function handleOrientation(event) {\n            const gamma = event.gamma - calibrationOffset.x;\n            const beta = event.beta - calibrationOffset.y;\n\n            const clampedGamma = Math.max(-45, Math.min(45, gamma));\n            const clampedBeta = Math.max(-45, Math.min(45, beta));\n\n            document.getElementById('angleX').textContent = gamma.toFixed(1) + '‹';\n            document.getElementById('angleY').textContent = beta.toFixed(1) + '‹';\n\n            const maxOffset = 140;\n            const bubbleX = (clampedGamma / 45) * maxOffset;\n            const bubbleY = (clampedBeta / 45) * maxOffset;\n\n            const bubble = document.getElementById('bubble');\n            bubble.style.transform = 	ranslate(calc(-50% + px), calc(-50% + px));\n\n            const totalTilt = Math.sqrt(gamma * gamma + beta * beta);\n            const tiltPercent = Math.min(100, (totalTilt / 45) * 100);\n            document.getElementById('tiltLevel').textContent = tiltPercent.toFixed(0) + '%';\n            document.getElementById('precisionFill').style.width = tiltPercent + '%';\n\n            if (soundEnabled) {\n                const absGamma = Math.abs(gamma);\n                const absBeta = Math.abs(beta);\n                const horizontalDominant = absGamma >= absBeta;\n\n                let freqBase;\n                let waveType;\n                let panValue = 0;\n\n                if (horizontalDominant) {\n                    freqBase = 420 + (gamma * 10);\n                    waveType = 'sawtooth';\n                    panValue = Math.max(-1, Math.min(1, gamma / 35));\n                } else {\n                    freqBase = 620 + (beta * 8);\n                    waveType = 'triangle';\n                }\n\n                const volume = Math.min(0.18, (totalTilt / 45) * 0.18);\n\n                if (totalTilt > 1) {\n                    createTone(freqBase, volume, { type: waveType, pan: panValue });\n                } else {\n                    stopContinuousTone();\n                }\n            }\n        }\n"
text = replace_function(text, 'handleOrientation', handle_block + '\n\n')

text = text.replace('            if (!continuousMode && oscillator) {\n                oscillator.stop();\n                oscillator = null;\n                gainNode = null;\n            }', '            if (!continuousMode && oscillator) {\n                oscillator.stop();\n                oscillator = null;\n                gainNode = null;\n                if (continuousPanner) {\n                    continuousPanner.disconnect();\n                    continuousPanner = null;\n                }\n            }')

HTML_PATH.write_text(text, encoding='utf-8')
