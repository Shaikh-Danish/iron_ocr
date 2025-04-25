import http from 'http';
import { spawn } from 'child_process';

const SUBDOMAIN = 'gcp-cloud-tunnel'; // Replace with your subdomain
const PORT = 5672; // The local port your app is running on
const CHECK_URL = `http://${SUBDOMAIN}.loca.lt`; // or tunnel URL like http://your-subdomain.loca.lt

function checkTunnel(callback) {
    const req = http.get(CHECK_URL, res => {
        callback(res.statusCode === 200);
    });

    req.on('error', () => callback(false));
    req.end();
}

function startTunnel() {
    console.log(`Starting localtunnel on subdomain: ${SUBDOMAIN}`);
    
    // If localtunnel is installed globally
    const lt = spawn('lt', ['--port', PORT, '--subdomain', SUBDOMAIN], {
        stdio: 'inherit',
        shell: true
    });

    lt.on('error', (err) => {
        console.error('Failed to start localtunnel:', err);
    });

    lt.on('exit', (code) => {
        console.log(`LocalTunnel exited with code ${code}`);
    });
}

checkTunnel((isRunning) => {
    if (isRunning) {
        console.log(`✅ LocalTunnel is already running at ${CHECK_URL}`);
    } else {
        console.log(`❌ LocalTunnel not running. Restarting...`);
        startTunnel();
    }
});
