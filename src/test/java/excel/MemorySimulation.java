package excel;

import java.util.Random;

public class MemorySimulation {
    private int pageSize;
    private int memorySize;
    private int[] memory;
    private boolean[] pageTable;
    private int pageFaultCount;

    public MemorySimulation(int pageSize, int memorySize) {
        this.pageSize = pageSize;
        this.memorySize = memorySize;
        this.memory = new int[memorySize];
        this.pageTable = new boolean[memorySize / pageSize];
        this.pageFaultCount = 0;
    }

    public void simulateMemoryAccess(int numAccesses) {
        Random random = new Random();

        for (int i = 0; i < numAccesses; i++) {
            int virtualAddress = random.nextInt(memorySize * 2); // 随机生成虚拟地址
            readMemory(virtualAddress);
        }

        System.out.println("内存利用率: " + calculateMemoryUtilization() + "%");
        System.out.println("缺页率: " + calculatePageFaultRate() + "%");
        System.out.println("页面置换开销: " + calculatePageReplacementOverhead());
    }

    private void readMemory(int virtualAddress) {
        int pageNumber = virtualAddress / pageSize;
        int offset = virtualAddress % pageSize;

        if (pageNumber >= pageTable.length) {
            System.out.println("访问越界：" + virtualAddress);
            return;
        }

        if (!pageTable[pageNumber]) {
            handlePageFault(pageNumber);
        }

        int physicalAddress = memory[pageNumber] + offset;
        System.out.println("读取内存地址：" + physicalAddress);
    }

    private void handlePageFault(int pageNumber) {
        pageTable[pageNumber] = true;
        memory[pageNumber] = generatePhysicalAddress(pageNumber);
        pageFaultCount++;
    }

    private int generatePhysicalAddress(int pageNumber) {
        // 在实际情况中，这里可能包含页面置换算法的实现
        // 这里简单地生成一个随机的物理地址
        Random random = new Random();
        return random.nextInt(memorySize);
    }

    private double calculateMemoryUtilization() {
        int usedPages = 0;
        for (boolean pagePresent : pageTable) {
            if (pagePresent) {
                usedPages++;
            }
        }

        return (double) usedPages / pageTable.length * 100;
    }

    private double calculatePageFaultRate() {
        return (double) pageFaultCount / (memorySize * 2) * 100;
    }

    private String calculatePageReplacementOverhead() {
        // 在实际情况中，这里可能需要更复杂的计算方式
        // 这里简单地返回页面置换次数
        return String.valueOf(pageFaultCount);
    }

    public static void main(String[] args) {
        int pageSize = 2048; // 页面大小为2k
        int memorySize = 8192; // 内存大小为8k
        int numAccesses = 2; // 模拟100次内存访问操作

        MemorySimulation memorySimulation = new MemorySimulation(pageSize, memorySize);
        memorySimulation.simulateMemoryAccess(numAccesses);
    }
}
