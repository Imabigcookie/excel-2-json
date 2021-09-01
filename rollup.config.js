import alias from '@rollup/plugin-alias'
import json from '@rollup/plugin-json'
import replace from '@rollup/plugin-replace'
import { nodeResolve } from '@rollup/plugin-node-resolve'
import commonjs from '@rollup/plugin-commonjs'
import image from '@rollup/plugin-image'
import postcss from 'rollup-plugin-postcss'
import vue from 'rollup-plugin-vue'
import esbuild from 'rollup-plugin-esbuild'
import serve from 'rollup-plugin-serve'
import livereload from 'rollup-plugin-livereload'
import filesize from 'rollup-plugin-filesize'
import requireContext from 'rollup-plugin-require-context'

const production = !process.env.ROLLUP_WATCH
const port = 8080

export default {
  input: 'src/main.js',
  output: {
    dir: 'public',
    entryFileNames: 'app.js',
    format: 'iife',
    sourcemap: !production ? 'inline' : false,
    name: 'app',
  },
  plugins: [
    json(),
    alias({
      entries: [{ find: '@', replacement: __dirname + '/src/' }],
    }),
    image(),
    postcss({ extract: 'app.css', config: { path: 'postcss.config.js' } }),
    requireContext(),
    nodeResolve({
      jsnext: true,
      main: true,
      browser: true,
    }),
    commonjs(),
    vue({ css: false, needMap: false }),
    replace({
      'process.env.NODE_ENV': production ? '"production"' : '"development"',
      preventAssignment: true,
    }),
    esbuild({
      minify: production,
      target: 'es2015',
    }),
    !production &&
      serve({
        open: true,
        contentBase: 'public',
        historyApiFallback: true,
        port,
      }),
    !production && livereload({ watch: 'public' }),
    production && filesize(),
  ],
  watch: {
    clearScreen: true,
  },
}
