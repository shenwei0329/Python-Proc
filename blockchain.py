# -*- coding: utf-8 -*-
#
#
#

from hashlib import sha256


class Block:

    def __init__(self, index, timestamp, data, previousHash=""):
        self.index = index
        self.timestamp = timestamp
        self.data = data
        self.previousHash = previousHash
        self.hash = self.calculateHash()

    def calculateHash(self):
        plainData = str(self.index) + str(self.timestamp) + str(self.data)
        return sha256(plainData.encode('utf-8')).hexdigest()

    def __str__(self):
        return str(self.__dict__)


class BlockChain:

    def __init__(self):
        self.chain = [self.createGenesisBlock()]

    def createGenesisBlock(self):
        return Block(0, "01/01/2018", "genesis block", "0")

    def getLatestBlock(self):
        return self.chain[len(self.chain) - 1]

    def addBlock(self, newBlock):
        newBlock.previousHash = self.getLatestBlock().hash
        newBlock.hash = newBlock.calculateHash()
        self.chain.append(newBlock)

    def __str__(self):
        return str(self.__dict__)

    def chainIsValid(self):
        for index in range(1, len(self.chain)):
            currentBlock = self.chain[index]
            previousBlock = self.chain[index - 1]
            if (currentBlock.hash != currentBlock.calculateHash()):
                return False
            if previousBlock.hash != currentBlock.previousHash:
                return False
        return True


myCoin = BlockChain()
myCoin.addBlock(Block(1, "02/01/2018", "{amount:4}"))
myCoin.addBlock(Block(2, "03/01/2018", "{amount:5}"))

# print block info
print("print block info ####:")
for block in myCoin.chain:
    print(block)
# check blockchain is valid
print("before tamper block,blockchain is valid ###")
print(myCoin.chainIsValid())
# tamper the blockinfo
myCoin.chain[1].data = "{amount:1002}"
print("after tamper block,blockchain is valid ###")
print(myCoin.chainIsValid())