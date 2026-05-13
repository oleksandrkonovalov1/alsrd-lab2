export class ArrayQueue {
  private data: number[];
  private head: number;
  private tail: number;
  private size: number;
  private capacity: number;

  constructor(capacity: number = 100) {
    this.data = new Array(capacity);
    this.head = 0;
    this.tail = 0;
    this.size = 0;
    this.capacity = capacity;
  }

  enqueue(value: number): void {
    if (this.size >= this.capacity) throw new Error("Queue overflow");
    this.data[this.tail] = value;
    this.tail = (this.tail + 1) % this.capacity;
    this.size++;
  }

  dequeue(): number {
    if (this.size === 0) throw new Error("Queue is empty");
    const value = this.data[this.head];
    this.head = (this.head + 1) % this.capacity;
    this.size--;
    return value;
  }

  isEmpty(): boolean {
    return this.size === 0;
  }

  print(): void {
    const items: number[] = [];
    for (let i = 0; i < this.size; i++) {
      items.push(this.data[(this.head + i) % this.capacity]);
    }
    console.log(`Queue [head → tail]: ${items.join(", ")}`);
  }
}
