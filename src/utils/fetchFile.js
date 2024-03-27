export async function fetchFile({ fileUrl }) {
    try {
        const response = await fetch(fileUrl);

        if (!response.ok) {
            const message = `An error has occured: ${response.status}`;
            throw new Error(message);
        }

        return response.blob();
    } catch (error) {
        console.error(error);
    }
}