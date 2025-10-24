import path from 'node:path';
import resolve from '@rollup/plugin-node-resolve';
import commonjs from '@rollup/plugin-commonjs';
import alias from '@rollup/plugin-alias';
import terser from '@rollup/plugin-terser';

const isProd = process.env.NODE_ENV === 'production';

export default {
    input: 'build-src/static/wrappers/copilotstudio-global.js',
    output: {
        file: 'static-resources-build/copilotStudioClient.js',
        format: 'iife',              // single global bundle
        name: 'Ignored',             // wrapper sets window.MicrosoftAgents directly
        sourcemap: false
    },
    plugins: [
        alias({
            entries: [
                {
                    find: 'math-random',
                    replacement: path.resolve('build-src/static/shims/math-random.js')
                }
            ]
        }),
        resolve({
            browser: true,
            preferBuiltins: false,
            mainFields: ['browser', 'module', 'main']
        }),
        commonjs(),
        terser({ format: { comments: false } }) // always minify for size
    ],
    treeshake: {
        moduleSideEffects: false
    }
};