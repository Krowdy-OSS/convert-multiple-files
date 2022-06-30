export declare const convertWordFiles: (pathFile: string, extOutput: string, outputDir: string) => Promise<string>;
export declare const convertWordFileToHTML: (pathFile: string, outputDir: string, outputPrefix: string) => Promise<{
    output: string;
}>;
export declare const convertToBase64: (pathFile: string) => Promise<string>;
