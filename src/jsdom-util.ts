import { JSDOM } from "jsdom"

export async function JSDOMfromURL(url: string, retries: number = 10) {
    let sleep = 1;
    while (retries > 0) {
        try {
            return await JSDOM.fromURL(url)
        }
        catch (e) {
            retries--;
            if (retries === 0) break;
            console.error(`Failed to fetch ${url}, retrying in ${sleep}s...`);
            await new Promise(r => setTimeout(r, sleep * 1000));
            sleep *= 2;
        }
    }
    throw new Error(`Failed to fetch ${url}`);
}