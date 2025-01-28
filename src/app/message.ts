import * as whatsapp from 'whatsapp-chat-parser';

export interface Message extends whatsapp.Message {
    in: number | '';
    out: number | '';
    net: number | '';
    cards: string;
    audrey: string;
    freeplay: string;
    w2g: string;
    multipleInOut: string;
}