export class ArrayStack {
  private data: number[];
  private top: number;
  private capacity: number;

  constructor(capacity: number = 100) {
    this.data = new Array(capacity);
    this.top = -1;
    this.capacity = capacity;
  }

  push(value: number): void {
    if (this.top >= this.capacity - 1) throw new Error("Stack overflow");
    this.data[++this.top] = value;
  }

  pop(): number {
    if (this.top < 0) throw new Error("Stack underflow");
    return this.data[this.top--];
  }

  peek(): number {
    if (this.top < 0) throw new Error("Stack is empty");
    return this.data[this.top];
  }

  isEmpty(): boolean {
    return this.top < 0;
  }

  print(): void {
    const items: number[] = [];
    for (let i = 0; i <= this.top; i++) items.push(this.data[i]);
    console.log(`Stack [bottom → top]: ${items.join(", ")}`);
  }
}
