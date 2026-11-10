/**
 * SICOF 3.0 PROTOTYPE — Motion Background Engine (WebGL Fluid Mesh Shader)
 * GPU-Accelerated a 60fps com Simplex Noise, amortecimento inercial e auto-pausa.
 */

class MotionBackground {
  constructor(canvasId) {
    this.canvas = document.getElementById(canvasId);
    if (!this.canvas) return;

    this.gl = this.canvas.getContext("webgl") || this.canvas.getContext("experimental-webgl");
    this.isReducedMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
    this.isActive = !this.isReducedMotion;
    this.animationFrameId = null;

    this.mouse = { x: 0.5, y: 0.5, targetX: 0.5, targetY: 0.5 };
    this.startTime = performance.now();

    // Cor dinâmica da Organização Militar ativa (Default Cyan)
    this.targetAccentR = 0.22;
    this.targetAccentG = 0.74;
    this.targetAccentB = 0.97;
    this.currentAccentR = 0.22;
    this.currentAccentG = 0.74;
    this.currentAccentB = 0.97;

    if (this.gl && this.isActive) {
      this.initGL();
      this.bindEvents();
      this.start();
    } else {
      this.fallbackCSS();
    }
  }

  initGL() {
    const gl = this.gl;

    const vsSource = `
      attribute vec2 a_position;
      void main() {
        gl_Position = vec4(a_position, 0.0, 1.0);
      }
    `;

    const fsSource = `
      precision mediump float;
      uniform float u_time;
      uniform vec2 u_resolution;
      uniform vec2 u_mouse;
      uniform vec3 u_accent_color;

      // Pseudo Simplex Noise 2D
      vec3 mod289(vec3 x) { return x - floor(x * (1.0 / 289.0)) * 289.0; }
      vec2 mod289(vec2 x) { return x - floor(x * (1.0 / 289.0)) * 289.0; }
      vec3 permute(vec3 x) { return mod289(((x*34.0)+1.0)*x); }

      float snoise(vec2 v) {
        const vec4 C = vec4(0.211324865405187, 0.366025403784439, -0.577350269189626, 0.024390243902439);
        vec2 i  = floor(v + dot(v, C.yy));
        vec2 x0 = v - i + dot(i, C.xx);
        vec2 i1 = (x0.x > x0.y) ? vec2(1.0, 0.0) : vec2(0.0, 1.0);
        vec4 x12 = x0.xyxy + C.xxzz;
        x12.xy -= i1;
        i = mod289(i);
        vec3 p = permute(permute(i.y + vec3(0.0, i1.y, 1.0)) + i.x + vec3(0.0, i1.x, 1.0));
        vec3 m = max(0.5 - vec3(dot(x0,x0), dot(x12.xy,x12.xy), dot(x12.zw,x12.zw)), 0.0);
        m = m*m; m = m*m;
        vec3 x = 2.0 * fract(p * C.www) - 1.0;
        vec3 h = abs(x) - 0.5;
        vec3 ox = floor(x + 0.5);
        vec3 a0 = x - ox;
        m *= 1.79284291400159 - 0.85373472095314 * (a0*a0 + h*h);
        vec3 g;
        g.x  = a0.x  * x0.x  + h.x  * x0.y;
        g.yz = a0.yz * x12.xz + h.yz * x12.yw;
        return 130.0 * dot(m, g);
      }

      void main() {
        vec2 st = gl_FragCoord.xy / u_resolution.xy;
        vec2 mouseOffset = (u_mouse - 0.5) * 0.15;
        
        // Ondulações fluidas multidimensionais
        float t = u_time * 0.08;
        float n1 = snoise(st * 1.8 + vec2(t * 0.5, t * 0.3) + mouseOffset);
        float n2 = snoise(st * 2.6 - vec2(t * 0.4, -t * 0.2) + n1 * 0.3);
        
        // Paleta de Alta Sofisticação Institucional
        vec3 colBase   = vec3(0.031, 0.043, 0.078); // Obsidian Slate #080b14
        vec3 colNavy   = vec3(0.058, 0.109, 0.211); // Deep Navy #0f1c36
        vec3 colTeal   = vec3(0.023, 0.250, 0.231); // Dark Emerald/Teal

        vec3 color = mix(colBase, colNavy, smoothstep(-0.6, 0.7, n1));
        color = mix(color, colTeal, smoothstep(-0.2, 0.8, n2) * 0.45);
        color += u_accent_color * (smoothstep(0.65, 0.98, n1 + n2 * 0.5) * 0.25);

        // Vinheta suave nas bordas
        float vignette = 1.0 - length((st - 0.5) * 1.3);
        color *= clamp(vignette, 0.4, 1.0);

        gl_FragColor = vec4(color, 1.0);
      }
    `;

    const vs = this.createShader(gl.VERTEX_SHADER, vsSource);
    const fs = this.createShader(gl.FRAGMENT_SHADER, fsSource);
    this.program = gl.createProgram();
    gl.attachShader(this.program, vs);
    gl.attachShader(this.program, fs);
    gl.linkProgram(this.program);

    if (!gl.getProgramParameter(this.program, gl.LINK_STATUS)) {
      console.warn("Falha no link do Shader WebGL:", gl.getProgramInfoLog(this.program));
      this.fallbackCSS();
      return;
    }

    // Geometria em tela cheia (quadrado de 2 triângulos)
    const positionBuffer = gl.createBuffer();
    gl.bindBuffer(gl.ARRAY_BUFFER, positionBuffer);
    gl.bufferData(
      gl.ARRAY_BUFFER,
      new Float32Array([-1, -1, 1, -1, -1, 1, -1, 1, 1, -1, 1, 1]),
      gl.STATIC_DRAW
    );

    this.positionLocation = gl.getAttribLocation(this.program, "a_position");
    this.timeLocation = gl.getUniformLocation(this.program, "u_time");
    this.resolutionLocation = gl.getUniformLocation(this.program, "u_resolution");
    this.mouseLocation = gl.getUniformLocation(this.program, "u_mouse");
    this.accentLocation = gl.getUniformLocation(this.program, "u_accent_color");

    this.resize();
  }

  createShader(type, source) {
    const gl = this.gl;
    const shader = gl.createShader(type);
    gl.shaderSource(shader, source);
    gl.compileShader(shader);
    if (!gl.getShaderParameter(shader, gl.COMPILE_STATUS)) {
      console.warn("Erro ao compilar shader:", gl.getShaderInfoLog(shader));
      gl.deleteShader(shader);
      return null;
    }
    return shader;
  }

  resize() {
    if (!this.canvas || !this.gl) return;
    const dpr = Math.min(window.devicePixelRatio || 1, 2);
    const width = window.innerWidth;
    const height = window.innerHeight;

    if (this.canvas.width !== width * dpr || this.canvas.height !== height * dpr) {
      this.canvas.width = width * dpr;
      this.canvas.height = height * dpr;
      this.gl.viewport(0, 0, this.canvas.width, this.canvas.height);
    }
  }

  bindEvents() {
    window.addEventListener("resize", () => this.resize(), { passive: true });

    window.addEventListener(
      "mousemove",
      (e) => {
        this.mouse.targetX = e.clientX / window.innerWidth;
        this.mouse.targetY = 1.0 - e.clientY / window.innerHeight;
      },
      { passive: true }
    );

    // Pausa automática quando a aba for minimizada ou estiver em segundo plano
    document.addEventListener("visibilitychange", () => {
      if (document.hidden) {
        this.stop();
      } else if (this.isActive) {
        this.start();
      }
    });

    // Pausa com IntersectionObserver se fora do campo de visão
    const observer = new IntersectionObserver(
      (entries) => {
        entries.forEach((entry) => {
          if (entry.isIntersecting && this.isActive) {
            this.start();
          } else {
            this.stop();
          }
        });
      },
      { threshold: 0.05 }
    );
    observer.observe(this.canvas);
  }

  render() {
    if (!this.isActive || !this.gl) return;

    // Amortecimento inercial suave do mouse (Spring Lerp)
    this.mouse.x += (this.mouse.targetX - this.mouse.x) * 0.04;
    this.mouse.y += (this.mouse.targetY - this.mouse.y) * 0.04;

    const gl = this.gl;
    const currentTime = (performance.now() - this.startTime) * 0.001;

    gl.useProgram(this.program);
    gl.enableVertexAttribArray(this.positionLocation);
    gl.vertexAttribPointer(this.positionLocation, 2, gl.FLOAT, false, 0, 0);

    gl.uniform1f(this.timeLocation, currentTime);
    gl.uniform2f(this.resolutionLocation, this.canvas.width, this.canvas.height);
    gl.uniform2f(this.mouseLocation, this.mouse.x, this.mouse.y);

    // Suavização e envio da cor de destaque da OM
    this.currentAccentR += (this.targetAccentR - this.currentAccentR) * 0.05;
    this.currentAccentG += (this.targetAccentG - this.currentAccentG) * 0.05;
    this.currentAccentB += (this.targetAccentB - this.currentAccentB) * 0.05;

    if (this.accentLocation) {
      gl.uniform3f(this.accentLocation, this.currentAccentR, this.currentAccentG, this.currentAccentB);
    }

    gl.drawArrays(gl.TRIANGLES, 0, 6);

    this.animationFrameId = requestAnimationFrame(() => this.render());
  }

  start() {
    if (!this.animationFrameId && this.isActive) {
      this.animationFrameId = requestAnimationFrame(() => this.render());
    }
  }

  stop() {
    if (this.animationFrameId) {
      cancelAnimationFrame(this.animationFrameId);
      this.animationFrameId = null;
    }
  }

  toggle() {
    this.isActive = !this.isActive;
    if (this.isActive) {
      this.canvas.style.opacity = "1";
      this.start();
    } else {
      this.stop();
      this.canvas.style.opacity = "0";
    }
    return this.isActive;
  }

  setAccentColor(hexColor) {
    if (!hexColor) return;
    let c = hexColor.replace("#", "");
    if (c.length === 3) c = c.split("").map((x) => x + x).join("");
    const num = parseInt(c, 16);
    if (isNaN(num)) return;
    this.targetAccentR = ((num >> 16) & 255) / 255;
    this.targetAccentG = ((num >> 8) & 255) / 255;
    this.targetAccentB = (num & 255) / 255;
  }

  fallbackCSS() {
    if (this.canvas) {
      this.canvas.style.display = "none";
    }
    document.body.classList.add("bg-fallback-mesh");
  }
}

window.MotionBackground = MotionBackground;
